using System;
using System.Runtime.InteropServices;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Scoped-блок оптимизации производительности (п.1.4 ТЗ).
    /// При входе отключает ScreenUpdating, EnableEvents и переводит Calculation
    /// в ручной режим; при выходе (в finally внутри Dispose) восстанавливает исходные значения.
    /// Поддерживает вложенность: каждый блок восстанавливает ровно то, что застал сам.
    /// </summary>
    /// <example>
    /// <code>
    /// using (app.BeginPerformanceScope())
    /// {
    ///     // массовые операции с книгой
    /// }
    /// // настройки Application восстановлены
    /// </code>
    /// </example>
    public sealed class PerformanceScope : IDisposable
    {
        private const int XlCalculationManual = -4142;
        private const int XlCalculationAutomatic = -4105;

        private readonly object _application;
        private readonly object? _oldScreenUpdating;
        private readonly object? _oldEnableEvents;
        private readonly object? _oldCalculation;
        private bool _disposed;

        internal PerformanceScope(ComExcelApplication app)
        {
            _application = app.Application!;
            // Запоминаем текущие значения, чтобы восстановить именно их
            _oldScreenUpdating = ComExcelApplication.GetAppProperty(_application, "ScreenUpdating");
            _oldEnableEvents = ComExcelApplication.GetAppProperty(_application, "EnableEvents");
            _oldCalculation = ComExcelApplication.GetAppProperty(_application, "Calculation");

            ComExcelApplication.SetAppProperty(_application, "ScreenUpdating", false);
            ComExcelApplication.SetAppProperty(_application, "EnableEvents", false);
            ComExcelApplication.SetAppProperty(_application, "Calculation", XlCalculationManual);
        }

        /// <summary>Восстанавливает настройки Excel (идемпотентно; вызывается в finally).</summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                ComExcelApplication.SetAppProperty(_application, "ScreenUpdating", _oldScreenUpdating ?? true);
                ComExcelApplication.SetAppProperty(_application, "EnableEvents", _oldEnableEvents ?? true);
                ComExcelApplication.SetAppProperty(_application, "Calculation", _oldCalculation ?? XlCalculationAutomatic);
            }
            catch
            {
                // Приложение могло быть уже освобождено — молча выходим
            }
        }
    }

    /// <summary>
    /// Лёгкая обёртка над COM-объектом Workbook (late binding).
    /// Освобождение RCW выполняет <see cref="ComExcelApplication"/> (список отслеживания),
    /// <see cref="Close"/> лишь сохраняет ссылку на «книгу закрыта».
    /// </summary>
    public sealed class ComWorkbook : IDisposable
    {
        private readonly ComExcelApplication _app;

        /// <summary>Raw COM-объект книги (доступен для внутреннего использования верхних слоёв).</summary>
        public object Raw { get; }

        /// <summary>Имя файла книги (без пути) или пустая строка для ещё не сохранённой.</summary>
        public string Name => ((dynamic)Raw).Name ?? string.Empty;

        /// <summary>Полный путь к файлу книги.</summary>
        public string FullName => ((dynamic)Raw).FullName ?? string.Empty;

        internal ComWorkbook(ComExcelApplication app, object raw)
        {
            _app = app;
            Raw = raw;
            app.Tracked(raw);
        }

        /// <summary>Коллекция имён книги (Names).</summary>
        public dynamic Names => ((dynamic)Raw).Names;

        /// <summary>Листы книги (Worksheets).</summary>
        public dynamic Worksheets => ((dynamic)Raw).Worksheets;

        /// <summary>Активный лист книги.</summary>
        public ComWorksheet ActiveSheet => _app.Tracked(new ComWorksheet(_app, (object)((dynamic)Raw).ActiveSheet));

        /// <summary>
        /// Закрывает книгу.
        /// </summary>
        /// <param name="saveChanges">Сохранять ли изменения в текущий файл.</param>
        public void Close(bool saveChanges = false)
        {
            try { ((dynamic)Raw).Close(SaveChanges: saveChanges); }
            catch (COMException) { /* книга могла быть уже закрыта */ }
        }

        /// <summary>Сохраняет книгу в текущий файл (Save).</summary>
        public void Save()
        {
            ((dynamic)Raw).Save();
        }

        /// <summary>
        /// Сохраняет книгу как файл. Формат выводится из расширения:
        /// .xlsx → 51 (xlOpenXMLWorkbook), .xlsm → 52, .xls → 56 (xlExcel8).
        /// </summary>
        /// <param name="path">Целевой путь.</param>
        /// <param name="fileFormat">Неявно задаваемый код формата; null — определить по расширению.</param>
        public void SaveAs(string path, int? fileFormat = null)
        {
            int format = fileFormat ?? GuessFormat(path);
            ((dynamic)Raw).SaveAs(path, FileFormat: format);
        }

        /// <summary>Определяет код xlFileFormat по расширению файла (.xlsx/.xlsm/.xls).</summary>
        /// <param name="path">Путь к файлу.</param>
        /// <exception cref="ArgumentException">Неизвестное расширение (в т.ч. .pdf — он выгружается через ExportAsFixedFormat).</exception>
        public static int GuessFormat(string path)
        {
            string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
            switch (ext)
            {
                case ".xlsx": return 51;   // xlOpenXMLWorkbook
                case ".xlsm": return 52;   // xlOpenXMLWorkbookMacroEnabled
                case ".xls": return 56;    // xlExcel8
                default:
                    throw new ArgumentException($"Неизвестное расширение файла Excel: '{ext}'. Ожидается .xlsx, .xlsm или .xls.");
            }
        }

        /// <summary>Закрывает книгу без сохранения и освобождает её COM-объект.</summary>
        public void Dispose()
        {
            try { Close(false); } catch { }
            _app.ReleaseTracked(Raw);
        }
    }

    /// <summary>
    /// Обёртка над COM-объектом Worksheet (late binding): чтение имени, видимости,
    /// копирование листов, активация, доступ к диапазонам.
    /// </summary>
    public sealed class ComWorksheet
    {
        private readonly ComExcelApplication _app;

        /// <summary>Raw COM-объект листа.</summary>
        public object Raw { get; }

        internal ComWorksheet(ComExcelApplication app, object raw)
        {
            _app = app;
            Raw = raw;
            app.Tracked(raw);
        }

        /// <summary>Имя листа.</summary>
        public string Name
        {
            get => (string)((dynamic)Raw).Name;
            set => ((dynamic)Raw).Name = value;
        }

        /// <summary>Видимость листа (XlSheetVisibility: -1 visible, 0 hidden, 2 veryHidden).</summary>
        public int Visible
        {
            get => (int)((dynamic)Raw).Visible;
            set => ((dynamic)Raw).Visible = value;
        }

        /// <summary>Индекс листа в книге (1-based).</summary>
        public int Index => (int)((dynamic)Raw).Index;

        /// <summary>Активировать лист.</summary>
        public void Activate()
        {
            ((dynamic)Raw).Activate();
        }

        /// <summary>Диапазон по адресу A1.</summary>
        /// <param name="a1Address">Адрес (например "A11:D38").</param>
        public ComRange Range(string a1Address) => _app.GetRange(this, a1Address);

        /// <summary>Ячейка по адресу A1.</summary>
        /// <param name="a1Address">Адрес ячейки.</param>
        public ComRange Cell(string a1Address) => _app.GetRange(this, a1Address);

        /// <summary>Используемая область листа (UsedRange).</summary>
        public ComRange UsedRange => _app.Tracked(new ComRange(_app, (object)((dynamic)Raw).UsedRange));

        /// <summary>
        /// Копирует лист перед указанным листом (или в конец, если target = null).
        /// Используется пагинацией (Этап 9): Worksheet.Copy(Before:=...).
        /// </summary>
        /// <param name="before">Лист-позиционир, перед которым вставляется копия; null — в конец.</param>
        /// <returns>Скопированный лист.</returns>
        public ComWorksheet CopyTo(ComWorksheet? before)
        {
            if (before == null)
            {
                ((dynamic)Raw).Copy(After: ((dynamic)Raw));
                // после копирования «сам за собой» новый лист — следующий по индексу
                dynamic sheets = ((dynamic)Raw).Parent.Worksheets;
                int newIndex = (int)((dynamic)Raw).Index + 1;
                return _app.Tracked(new ComWorksheet(_app, (object)sheets[newIndex]));
            }
            else
            {
                ((dynamic)Raw).Copy(Before: before.Raw);
                dynamic sheets = ((dynamic)Raw).Parent.Worksheets;
                int copyIndex = before.Index; // копия встала на место target, сдвинув его
                return _app.Tracked(new ComWorksheet(_app, (object)sheets[copyIndex]));
            }
        }

        /// <summary>Удаляет лист из книги (DisplayAlerts подавлен на уровне Application).</summary>
        public void Delete()
        {
            ((dynamic)Raw).Delete();
        }
    }

    /// <summary>
    /// Обёртка над COM-объектом Range: пакетная запись/чтение значений через один вызов
    /// <c>Range.Value2</c> (п.1.8), формулы, адрес, объединённые ячейки.
    /// Массивы передаются как object[,] (0-based), Value2 принимает и возвращает их напрямую.
    /// </summary>
    public sealed class ComRange
    {
        private readonly ComExcelApplication _app;

        /// <summary>Raw COM-объект диапазона.</summary>
        public object Raw { get; }

        internal ComRange(ComExcelApplication app, object raw)
        {
            _app = app;
            Raw = raw;
            app.Tracked(raw);
        }

        /// <summary>Адрес диапазона в нотации A1 (относительно листа).</summary>
        public string Address => (string)((dynamic)Raw).Address[true, true];

        /// <summary>Номер первой строки диапазона.</summary>
        public int Row => (int)((dynamic)Raw).Row;

        /// <summary>Номер первого столбца диапазона.</summary>
        public int Column => (int)((dynamic)Raw).Column;

        /// <summary>Количество строк диапазона.</summary>
        public int Rows => (int)((dynamic)Raw).Rows.Count;

        /// <summary>Количество столбцов диапазона.</summary>
        public int Columns => (int)((dynamic)Raw).Columns.Count;

        /// <summary>Формула верхней левой ячейки (или null, если константа).</summary>
        public string? Formula
        {
            get
            {
                object? f = ((dynamic)Raw).Formula;
                return f as string;
            }
            set => ((dynamic)Raw).Formula = value;
        }

        /// <summary>
        /// Пакетная запись значений: одно обращение к COM через <c>Value2</c> со массивом object[,].
        /// Массив должен соответствовать размеру диапазона (строки × столбцы).
        /// </summary>
        /// <param name="values">Двумерный массив значений (0-based).</param>
        /// <exception cref="ExcelRangeOverflowException">Размер массива больше диапазона.</exception>
        public void SetValue2(object[,] values)
        {
            if (values.GetLength(0) > Rows || values.GetLength(1) > Columns)
                throw new ExcelRangeOverflowException(
                    $"Массив данных {values.GetLength(0)}x{values.GetLength(1)} не помещается в диапазон {Address} ({Rows}x{Columns}).",
                    values.GetLength(0) * values.GetLength(1),
                    Rows * Columns);

            ((dynamic)Raw).Value2 = values;
        }

        /// <summary>Запись одного значения в ячейку/диапазон.</summary>
        /// <param name="value">Значение.</param>
        public void SetValue(object? value)
        {
            ((dynamic)Raw).Value2 = value;
        }

        /// <summary>Чтение всех значений диапазона одним вызовом Value2.</summary>
        /// <returns>object[,] для многоячеечного диапазона или одиночное значение для одной ячейки.</returns>
        public object? GetValue2()
        {
            return ((dynamic)Raw).Value2;
        }

        /// <summary>Чтение значений диапазона в нормализованный object[,] (строки × столбцы).</summary>
        /// <returns>Двумерный массив (0-based).</returns>
        public object?[,] ReadAsMatrix()
        {
            object? raw = GetValue2();
            int rows = Rows, cols = Columns;
            var result = new object?[rows, cols];

            if (raw is object[,] arr) // многоячеечный диапазон
            {
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        result[r, c] = arr[r + 1, c + 1]; // Value2 возвращает 1-based массив
            }
            else
            {
                result[0, 0] = raw;
            }
            return result;
        }

        /// <summary>Печать области листа (PageSetup.PrintArea = ...).</summary>
        /// <param name="sheet">Лист, которому назначается PrintArea.</param>
        public void SetPrintAreaOf(ComWorksheet sheet)
        {
            ((dynamic)sheet.Raw).PageSetup.PrintArea = Address;
        }

        /// <summary>Колонтитулы диапазона (Header/Footer через PageSetup листа).</summary>
        /// <param name="sheet">Лист.</param>
        /// <param name="centerHeader">Центр верхнего колонтитула.</param>
        /// <param name="centerFooter">Центр нижнего колонтитула.</param>
        public static void SetHeaders(ComWorksheet sheet, string? centerHeader, string? centerFooter)
        {
            dynamic ps = ((dynamic)sheet.Raw).PageSetup;
            if (centerHeader != null) ps.CenterHeader = centerHeader;
            if (centerFooter != null) ps.CenterFooter = centerFooter;
        }

        /// <summary>Перемещает/копирует содержимое диапазона в другой диапазон (значения+форматы без буфера обмена).</summary>
        /// <param name="target">Целевой диапазон.</param>
        public void CopyValuesTo(ComRange target)
        {
            ((dynamic)target.Raw).Value2 = ((dynamic)Raw).Value2;
            ((dynamic)target.Raw).NumberFormat = ((dynamic)Raw).NumberFormat;
        }
    }
}
