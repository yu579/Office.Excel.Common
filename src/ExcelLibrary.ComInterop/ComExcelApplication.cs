using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Режим запуска Excel: создать новый экземпляр или подключиться к уже запущенному.
    /// </summary>
    public enum ExcelApplicationMode
    {
        /// <summary>Создать новый скрытый экземпляр EXCEL.EXE и завершить его при Dispose.</summary>
        CreateNew,

        /// <summary>Подключиться к существующему запущенному экземпляру Excel (GetActiveObject);
        /// если не найден — создать новый. Чужой процесс при Dispose НЕ завершается.</summary>
        AttachToRunning,
    }

    /// <summary>
    /// Внутренняя COM-обёртка над <c>Excel.Application</c> (нижний слой библиотеки).
    /// Инкапсулирует создание/подключение, подавление диалогов, контроль битности,
    /// отслеживание PID запущенного процесса и гарантированное освобождение COM-объектов.
    /// Публичный API верхних слоёв не зависит от типов PIA: при недоступности
    /// Primary Interop Assembly используется late binding (COM-объекты как <c>System.__ComObject</c>).
    /// </summary>
    /// <remarks>
    /// Использование:
    /// <code>
    /// using var app = new ComExcelApplication();
    /// using (app.BeginPerformanceScope())
    /// {
    ///     var book = app.OpenWorkbook(path);
    ///     ...
    ///     book.Close(false);
    /// }
    /// </code>
    /// </remarks>
    public sealed class ComExcelApplication : IDisposable
    {
        private readonly List<object> _trackedComObjects = new List<object>();
        private bool _disposed;
        private int? _ownedProcessId; // PID только если процесс создан нами

        /// <summary>COM-объект Excel.Application (RCW из PIA или dynamic из late binding).</summary>
        public object? Application { get; private set; }

        /// <summary>Использулся ли late binding (PIA оказался недоступен).</summary>
        public bool UsingLateBinding { get; private set; }

        /// <summary>Режим запуска, выбранный при создании.</summary>
        public ExcelApplicationMode Mode { get; private set; }

        /// <summary>
        /// PID процесса EXCEL.EXE, созданного библиотекой; null — если подключение к чужому экземпляру
        /// или процесс не был определён.
        /// </summary>
        public int? OwnedProcessId => _ownedProcessId;

        /// <summary>True, если объект уже освобождён.</summary>
        public bool IsDisposed => _disposed;

        /// <summary>
        /// Создаёт (или подключается к) экземпляру Excel.Application с подавлением диалогов.
        /// </summary>
        /// <param name="mode">Режим запуска.</param>
        /// <param name="visible">Показывать ли окно Excel (по умолчанию — нет).</param>
        /// <param name="checkBitness">Проверять ли совпадение битности процесса и Excel.</param>
        /// <exception cref="ExcelBitnessMismatchException">Битность процесса и Excel различаются.</exception>
        /// <exception cref="ExcelLibraryException">Excel не установлен или COM-объект не создан.</exception>
        public ComExcelApplication(
            ExcelApplicationMode mode = ExcelApplicationMode.CreateNew,
            bool visible = false,
            bool checkBitness = true)
        {
            Mode = mode;

            if (checkBitness && Environment.OSVersion.Platform == PlatformID.Win32NT)
                EnsureBitnessMatches();

            object? application = null;
            switch (mode)
            {
                case ExcelApplicationMode.AttachToRunning:
                    if (!TryAttach(out application))
                        CreateInstance();
                    break;
                default:
                    CreateInstance();
                    break;
            }

            // TryAttach/CreateInstance записывают объект в поле Application
            if (Application == null)
                throw new ExcelLibraryException("Не удалось получить COM-объект Excel.Application.");
            Tracked(Application);

            // Подавление диалогов и фоновых запросов (п.1 чек-листа 1.1)
            try { ((dynamic)Application!).Visible = visible; } catch { /* некоторые сборки не позволяют до открытия */ }
            SetAppProperty("DisplayAlerts", false);
            SetAppProperty("AskToUpdateLinks", false);
            try { ((dynamic)Application!).AlertBeforeOverwrite = false; } catch { /* необязательное */ }
        }

        /// <summary>
        /// Подключается к активному экземпляру Excel через ROT (GetActiveObject).
        /// Найденный процесс помечается как «чужой» — при Dispose он не завершается (п.1.6).
        /// </summary>
        private bool TryAttach(out object application)
        {
            application = null!;
            try
            {
                // Marshal.GetActiveObject отсутствует в .NET 5+: используем oleaut32.coGetActiveObject
                int hr = CoGetActiveObject(ref _excelClsid, out object obj);
                if (hr == 0 && obj != null)
                {
                    Application = obj;
                    application = obj;
                    _ownedProcessId = null; // чужой экземпляр — не убиваем
                    return true;
                }
            }
            catch
            {
                // активный Excel не найден — вызывающий код создаст новый
            }
            return false;
        }

        // CLSID Excel.Application: {00024500-0000-0000-C000-000000000046}
        private static Guid _excelClsid = new Guid("00024500-0000-0000-C000-000000000046");

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void CoGetActiveObjectSafe([In] ref Guid clsid, [MarshalAs(UnmanagedType.IDispatch)] out object obj);

        private static int CoGetActiveObject(ref Guid clsid, out object obj)
        {
            try
            {
                CoGetActiveObjectSafe(ref clsid, out obj);
                return 0;
            }
            catch (COMException)
            {
                obj = null!;
                return unchecked((int)0x800401E3); // MK_E_UNAVAILABLE
            }
        }

        /// <summary>
        /// Создаёт новый экземпляр Excel: сначала через тип из PIA, при неудаче — late binding
        /// (Type.GetTypeFromProgID("Excel.Application")). Запоминает PID созданного процесса (п.1.5).
        /// </summary>
        private void CreateInstance()
        {
            var before = SnapshotExcelPids();

            // Попытка 1: раннее связывание через PIA (тип из Microsoft.Office.Interop.Excel).
            try
            {
                var piaType = typeof(Microsoft.Office.Interop.Excel.ApplicationClass);
                object created = Activator.CreateInstance(piaType)!;
                Application = created;
                UsingLateBinding = false;
            }
            catch
            {
                // PIA недоступна или не зарегистрирована — fallback на late binding (п.1.2)
                Application = CreateViaLateBinding();
                UsingLateBinding = true;
            }

            // Определяем PID нового процесса: появившийся EXCEL.EXE после CreateInstance
            _ownedProcessId = DetectNewlyStartedPid(before);
        }

        /// <summary>Late-binding создание COM-объекта Excel (Type.GetTypeFromProgID + Activator).</summary>
        private static object CreateViaLateBinding()
        {
            var t = Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
            if (t == null)
                throw new ExcelLibraryException(
                    "Microsoft Excel не обнаружен в реестре (ProgID 'Excel.Application'). " +
                    "Установите Microsoft Excel для работы библиотеки через COM Interop.");

            try
            {
                return Activator.CreateInstance(t)!;
            }
            catch (Exception ex)
            {
                throw new ExcelLibraryException(
                    "Не удалось создать COM-объект Excel.Application. Убедитесь, что Microsoft Excel " +
                    "установлен и доступен для автоматизации (COM).", ex);
            }
        }

        /// <summary>Регистрирует COM-объект для гарантированного освобождения в Dispose.</summary>
        /// <typeparam name="T">Тип обёртки или raw COM-объекта.</typeparam>
        /// <param name="comObject">Объект, за которым нужно следить.</param>
        /// <returns>Тот же объект для удобства цепочек.</returns>
        public T Tracked<T>(T comObject)
        {
            if (!_disposed && comObject != null)
                _trackedComObjects.Add(comObject!);
            return comObject;
        }

        /// <summary>
        /// Открывает книгу. Все созданные в процессе книги COM-объекты оборачиваются
        /// в <see cref="ComWorkbook"/> и автоматически освобождаются.
        /// </summary>
        /// <param name="path">Полный путь к файлу .xls/.xlsx/.xlsm.</param>
        /// <param name="readOnly">Открыть только для чтения.</param>
        /// <returns>Обёртка книги.</returns>
        /// <exception cref="FileNotFoundException">Файл не найден.</exception>
        /// <exception cref="ExcelTemplateException">Книга не открылась.</exception>
        public ComWorkbook OpenWorkbook(string path, bool readOnly = false)
        {
            ThrowIfDisposed();
            if (!File.Exists(path))
                throw new FileNotFoundException("Файл Excel не найден.", path);

            try
            {
                dynamic books = ((dynamic)Application!).Workbooks;
                dynamic book = books.Open(path, UpdateLinks: 0, ReadOnly: readOnly);
                return Tracked(new ComWorkbook(this, (object)book));
            }
            catch (COMException ex)
            {
                throw new ExcelTemplateException($"Не удалось открыть книгу '{path}': {ex.Message}", ex, path);
            }
        }

        /// <summary>Создаёт новую пустую книгу (для тестов и временных документов).</summary>
        /// <returns>Обёртка новой книги.</returns>
        public ComWorkbook AddWorkbook()
        {
            ThrowIfDisposed();
            dynamic books = ((dynamic)Application!).Workbooks;
            dynamic book = books.Add();
            return Tracked(new ComWorkbook(this, (object)book));
        }

        /// <summary>Сохраняет книгу под новым именем (SaveAs).</summary>
        /// <param name="book">Книга (обёртка или raw COM-объект).</param>
        /// <param name="path">Целевой путь.</param>
        /// <param name="fileFormat">Код формата xlFileFormat (51 = xlOpenXMLWorkbook, 56 = xlExcel8).</param>
        public void SaveWorkbookAs(object book, string path, int fileFormat)
        {
            ThrowIfDisposed();
            ((dynamic)book).SaveAs(path, FileFormat: fileFormat);
        }

        /// <summary>Возвращает коллекцию листов книги в виде обёрток.</summary>
        /// <param name="book">Обёртка книги.</param>
        public List<ComWorksheet> GetWorksheets(ComWorkbook book)
        {
            ThrowIfDisposed();
            var result = new List<ComWorksheet>();
            dynamic sheets = ((dynamic)book.Raw).Worksheets;
            int count = (int)sheets.Count;
            for (int i = 1; i <= count; i++)
                result.Add(Tracked(new ComWorksheet(this, (object)sheets[i])));
            return result;
        }

        /// <summary>Находит лист по имени (без учёта регистра).</summary>
        /// <param name="book">Обёртка книги.</param>
        /// <param name="name">Имя листа.</param>
        /// <returns>Обёртка листа.</returns>
        /// <exception cref="ExcelTemplateException">Лист с таким именем отсутствует.</exception>
        public ComWorksheet GetWorksheet(ComWorkbook book, string name)
        {
            ThrowIfDisposed();
            foreach (var ws in GetWorksheets(book))
            {
                if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                    return ws;
                MarshalRelease((object)ws.Raw);
            }
            throw new ExcelTemplateException($"Лист '{name}' не найден в книге.", name);
        }

        /// <summary>Возвращает диапазон листа по адресу A1 (например, "A11:D38").</summary>
        /// <param name="sheet">Обёртка листа.</param>
        /// <param name="a1Address">Адрес в нотации A1.</param>
        /// <returns>Обёртка диапазона.</returns>
        public ComRange GetRange(ComWorksheet sheet, string a1Address)
        {
            ThrowIfDisposed();
            dynamic range = ((dynamic)sheet.Raw).Range[a1Address];
            return Tracked(new ComRange(this, (object)range));
        }

        /// <summary>Возвращает именованный диапазон книги.</summary>
        /// <param name="book">Обёртка книги.</param>
        /// <param name="name">Имя диапазона.</param>
        /// <returns>Обёртка диапазона.</returns>
        /// <exception cref="ExcelTemplateException">Именованный диапазон отсутствует.</exception>
        public ComRange GetNamedRange(ComWorkbook book, string name)
        {
            ThrowIfDisposed();
            try
            {
                dynamic names = ((dynamic)book.Raw).Names;
                dynamic named = names.Item(name);
                dynamic refRange = named.RefersToRange;
                return Tracked(new ComRange(this, (object)refRange));
            }
            catch (COMException ex)
            {
                throw new ExcelTemplateException($"Именованный диапазон '{name}' не найден в книге.", ex, name);
            }
        }

        /// <summary>
        /// Возвращает scoped-блок оптимизации производительности (п.1.4):
        /// ScreenUpdating=false, EnableEvents=false, Calculation=xlCalculationManual
        /// с восстановлением исходных значений в finally.
        /// </summary>
        /// <returns>IDisposable-блок; настройки восстанавливаются при Dispose.</returns>
        public PerformanceScope BeginPerformanceScope()
        {
            ThrowIfDisposed();
            return new PerformanceScope(this);
        }

        /// <summary>Значение свойства Application (через reflection/dynamic), без исключения при отсутствии.</summary>
        internal object? GetAppProperty(string propertyName)
            => Application != null ? GetAppProperty(Application, propertyName) : null;

        internal void SetAppProperty(string propertyName, object value)
        {
            if (Application != null) SetAppProperty(Application, propertyName, value);
        }

        /// <summary>Статический доступ к свойству COM-объекта Application (для PerformanceScope после Dispose).</summary>
        internal static object? GetAppProperty(object application, string propertyName)
        {
            try { return application.GetType().GetProperty(propertyName)?.GetValue(application); }
            catch { return null; }
        }

        /// <summary>Статическая установка свойства COM-объекта Application; ошибки игнорируются.</summary>
        internal static void SetAppProperty(object application, string propertyName, object value)
        {
            try { application.GetType().GetProperty(propertyName)?.SetValue(application, value); }
            catch { /* свойство могло стать недоступным в конкретной сборке Excel */ }
        }

        /// <summary>Освобождает один COM-объект (RCW) и удаляет его из списка отслеживания.</summary>
        internal void ReleaseTracked(object comObject)
        {
            _trackedComObjects.Remove(comObject);
            MarshalRelease(comObject);
        }

        private static void MarshalRelease(object o)
        {
            try { Marshal.FinalReleaseComObject(o); }
            catch { /* объект мог быть уже освобождён */ }
        }

        /// <summary>Проверяет, что объект жив; иначе — <see cref="ObjectDisposedException"/>.</summary>
        internal void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ComExcelApplication),
                    "COM-объект Excel.Application уже освобождён (Dispose выполнен).");
        }

        /// <summary>
        /// Идемпотентное освобождение всех ресурсов (п.1.5): восстановление настроек,
        /// Quit() для созданного нами экземпляра, FinalReleaseComObject по всем отслеженным
        /// объектам в обратном порядке. Чужие процессы (AttachToRunning) не завершаются.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return; // идемпотентность
            _disposed = true;

            try
            {
                if (Application != null)
                {
                    SetAppProperty("ScreenUpdating", true);
                    SetAppProperty("EnableEvents", true);
                    SetAppProperty("Calculation", -4105); // xlCalculationAutomatic

                    bool ownsProcess = Mode == ExcelApplicationMode.CreateNew || _ownedProcessId.HasValue;
                    if (ownsProcess)
                    {
                        try { ((dynamic)Application!).Quit(); } catch { /* процесс мог завершиться сам */ }
                    }
                }
            }
            catch { /* на этапе очистки не бросаем */ }
            finally
            {
                // Освобождаем все отслеженные COM-объекты в обратном порядке создания
                for (int i = _trackedComObjects.Count - 1; i >= 0; i--)
                {
                    try { MarshalRelease(_trackedComObjects[i]); } catch { }
                }
                _trackedComObjects.Clear();
                Application = null!;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        #region Битность и процессы

        /// <summary>
        /// Сверяет разрядность текущего процесса и установленного Excel
        /// (по пути Excel.exe в Program Files (x86)/Program Files). При несовпадении —
        /// <see cref="ExcelBitnessMismatchException"/> с инструкцией (п.1.3).
        /// </summary>
        private static void EnsureBitnessMatches()
        {
            int processBits = Environment.Is64BitProcess ? 64 : 32;
            int? excelBits = DetectInstalledExcelBitness();
            if (excelBits.HasValue && excelBits.Value != processBits)
                throw new ExcelBitnessMismatchException(processBits, excelBits);
        }

        /// <summary>Определяет битность установленного Excel по расположению Excel.exe в реестре/Program Files.</summary>
        private static int? DetectInstalledExcelBitness()
        {
            try
            {
                string? exePath = ResolveExcelExePath();
                if (exePath == null || !File.Exists(exePath)) return null;

                // 32-битный Office всегда ставится в "Program Files (x86)"
                string normalized = Path.GetFullPath(exePath).Replace('/', '\\');
                if (normalized.IndexOf("Program Files (x86)", StringComparison.OrdinalIgnoreCase) >= 0)
                    return 32;
                if (normalized.IndexOf("Program Files", StringComparison.OrdinalIgnoreCase) >= 0)
                    return 64;
            }
            catch { /* определение best-effort */ }
            return null;
        }

        private static string? ResolveExcelExePath()
        {
            try
            {
                var t = Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
                return t?.AssemblyQualifiedName != null
                    ? ReadLocalServerPath()
                    : null;
            }
            catch { return null; }
        }

        /// <summary>Читает путь LocalServer32 из HKCR\Excel.Application\CLSID (best effort).</summary>
        private static string? ReadLocalServerPath()
        {
            try
            {
                using var root = Microsoft.Win32.Registry.ClassesRoot;
                using var clsidKey = root.OpenSubKey("Excel.Application\\CLSID");
                var clsid = clsidKey?.GetValue(string.Empty) as string;
                if (string.IsNullOrEmpty(clsid)) return null;
                using var localServer = Microsoft.Win32.Registry.ClassesRoot
                    .OpenSubKey($"CLSID\\{clsid}\\LocalServer32");
                var cmdRaw = localServer?.GetValue(string.Empty) as string;
                if (string.IsNullOrEmpty(cmdRaw)) return null;
                string cmd = cmdRaw!.Trim();
                if (cmd.StartsWith("\""))
                {
                    int end = cmd.IndexOf('"', 1);
                    return end > 1 ? cmd.Substring(1, end - 1) : null;
                }
                int space = cmd.IndexOf(' ');
                return space > 0 ? cmd.Substring(0, space) : cmd;
            }
            catch { return null; }
        }

        private static HashSet<int> SnapshotExcelPids()
        {
            var set = new HashSet<int>();
            try
            {
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                    set.Add(p.Id);
            }
            catch { }
            return set;
        }

        private static int? DetectNewlyStartedPid(HashSet<int> before)
        {
            try
            {
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id))
                        return p.Id;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Проверяет, завершился ли процесс Excel, созданный библиотекой (используется тестами
        /// для критерия 10.1.1: отсутствие «висящих» EXCEL.EXE).
        /// </summary>
        /// <param name="pid">PID процесса.</param>
        /// <param name="timeoutMs">Максимальное ожидание завершения, мс.</param>
        /// <returns>True — процесс завершился в пределах ожидания.</returns>
        public static bool WaitForOwnedProcessExit(int pid, int timeoutMs = 5000)
        {
            try
            {
                using var p = Process.GetProcessById(pid);
                return p.WaitForExit(timeoutMs);
            }
            catch (ArgumentException)
            {
                return true; // процесс уже отсутствует
            }
            catch { return false; }
        }

        #endregion
    }
}
