using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Жизненный цикл документа (Этап 2, п.4.4 ТЗ): открытие шаблона (путь/Stream) либо новой книги,
    /// ведение скрытого временного рабочего файла рядом с целевым путём, SaveAs в видимый файл
    /// с форматом по расширению (.xls/.xlsx/.xlsm), удаление временного файла при любой ошибке
    /// до успешного сохранения.
    /// </summary>
    internal sealed class DocumentLifecycle : IDisposable
    {
        private const int XlExcel8 = 56;          // .xls
        private const int XlOpenXmlWorkbook = 51; // .xlsx
        private const int XlOpenXmlMacro = 52;    // .xlsm

        private readonly ComExcelApplication _app;
        private readonly string? _targetPath;   // целевой («видимый») путь из WithHiddenWorkFile
        private string? _tempFilePath;          // скрытый временный рабочий файл
        private bool _saved;                    // флаг успешного финального сохранения
        private bool _disposed;

        /// <summary>Текущая рабочая книга.</summary>
        public ComWorkbook Workbook { get; private set; } = null!;

        /// <summary>Целевой путь сохранения (null — сохранение не настроено).</summary>
        public string? TargetPath => _targetPath;

        /// <summary>Определяется ли формат по расширению целевого файла как .xls (лимит 65 536 строк).</summary>
        public bool IsXlsTarget =>
            _targetPath != null &&
            Path.GetExtension(_targetPath).Equals(".xls", StringComparison.OrdinalIgnoreCase);

        internal DocumentLifecycle(ComExcelApplication app, string? targetPath)
        {
            _app = app;
            _targetPath = targetPath == null ? null : Path.GetFullPath(targetPath);
        }

        /// <summary>Открыть новую пустую книгу без шаблона (п.2.3).</summary>
        public void OpenNewWorkbook()
        {
            Workbook = _app.AddWorkbook();
            if (_targetPath != null)
                MaterializeTempFromCurrentWorkbook();
        }

        /// <summary>
        /// Открыть шаблон из файла: копируем его в скрытый временный файл рядом с целевым путём
        /// и открываем копию — исходный шаблон остаётся нетронутым (п.4.4).
        /// </summary>
        /// <param name="templatePath">Путь к файлу шаблона.</param>
        /// <exception cref="FileNotFoundException">Шаблон не найден.</exception>
        /// <exception cref="ExcelTemplateException">Шаблон не открылся.</exception>
        public void OpenFromTemplate(string templatePath)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Файл шаблона не найден.", templatePath);

            string temp = CreateTempPath(nearTarget: _targetPath ?? templatePath, extension: Path.GetExtension(templatePath));
            try
            {
                File.Copy(templatePath, temp);
                Hide(temp);
                _tempFilePath = temp;
                Workbook = _app.OpenWorkbook(temp);
            }
            catch
            {
                TryDeleteTemp();
                throw;
            }
        }

        /// <summary>
        /// Открыть шаблон из потока: содержимое сохраняется во временный файл (рядом с целевым путём),
        /// затем открывается книга (п.2.2).
        /// </summary>
        /// <param name="stream">Поток с содержимым книги.</param>
        /// <exception cref="ExcelTemplateException">Формат потока неизвестен или книга не открылась.</exception>
        public void OpenFromStream(Stream stream)
        {
            string ext = GuessStreamExtension(stream);
            string temp = CreateTempPath(nearTarget: _targetPath ?? Path.GetTempPath(), extension: ext);
            try
            {
                using (var fs = new FileStream(temp, FileMode.Create, FileAccess.Write))
                {
                    if (stream.CanSeek) stream.Position = 0;
                    stream.CopyTo(fs);
                }
                Hide(temp);
                _tempFilePath = temp;
                Workbook = _app.OpenWorkbook(temp);
            }
            catch
            {
                TryDeleteTemp();
                throw;
            }
        }

        /// <summary>
        /// Сохранить рабочий файл в целевой путь через SaveAs с форматом по расширению (п.2.5).
        /// После успешного сохранения временный файл удаляется.
        /// </summary>
        /// <exception cref="InvalidOperationException">Целевой путь не задан (WithHiddenWorkFile не вызван).</exception>
        /// <exception cref="ExcelLibraryException">Сохранение не удалось — временный файл удалён.</exception>
        public void SaveToTarget()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(DocumentLifecycle));
            if (_targetPath == null)
                throw new InvalidOperationException(
                    "Целевой путь не задан: используйте WithHiddenWorkFile(path) или SaveTo(path).");
            SaveTo(_targetPath);
        }

        /// <summary>Сохранить в произвольный путь с определением формата по расширению.</summary>
        /// <param name="path">Целевой путь (.xls/.xlsx/.xlsm).</param>
        public void SaveTo(string path)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(DocumentLifecycle));
            string full = Path.GetFullPath(path);
            int format = GuessFormatByExtension(full);

            string? dir = Path.GetDirectoryName(full);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            try
            {
                Workbook.SaveAs(full, format);
                _saved = true;
                TryDeleteTemp(); // после успешного сохранения временный файл больше не нужен
            }
            catch (Exception ex)
            {
                // п.4.4 / критерий 10.1.3: при ошибке записи временный файл удаляется, целевой не считается созданным
                TryDeleteTemp();
                throw new ExcelLibraryException($"Не удалось сохранить документ в '{full}'. Временный файл удалён.", ex);
            }
        }

        /// <summary>Проверка лимита строк для .xls (п.2.7, дефолт по открытому вопросу 4).</summary>
        /// <param name="startRow">Номер первой строки записи.</param>
        /// <param name="rowCount">Количество записываемых строк.</param>
        /// <exception cref="ExcelRangeOverflowException">Запись выходит за 65 536 строк листа .xls.</exception>
        public void EnsureWithinSheetLimit(int startRow, int rowCount)
        {
            int maxRows = 1_048_576; // .xlsx/.xlsm
            if (IsXlsTarget || (Workbook != null && Workbook.FullName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
                                && !Workbook.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                                && !Workbook.FullName.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)))
            {
                maxRows = 65_536;
            }

            long lastRow = (long)startRow + rowCount - 1;
            if (lastRow > maxRows)
                throw new ExcelRangeOverflowException(
                    $"Данные ({rowCount} строк, начиная со строки {startRow}) выходят за лимит {maxRows} строк листа" +
                    (maxRows == 65_536 ? " формата .xls." : "."),
                    requestedCount: (int)Math.Min(lastRow, int.MaxValue),
                    maximumCount: maxRows);
        }

        /// <summary>Код xlFileFormat по расширению пути (.xls=56, .xlsx=51, .xlsm=52).</summary>
        /// <param name="path">Путь к файлу.</param>
        /// <exception cref="ArgumentException">Неизвестное расширение.</exception>
        public static int GuessFormatByExtension(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            switch (ext)
            {
                case ".xls": return XlExcel8;
                case ".xlsx": return XlOpenXmlWorkbook;
                case ".xlsm": return XlOpenXmlMacro;
                default:
                    throw new ArgumentException($"Неизвестное расширение '{ext}': ожидаются .xls, .xlsx или .xlsm.");
            }
        }

        /// <summary>Формат вывода по расширению (для публичной информации/тестов).</summary>
        public static ExcelOutputFormat GetOutputFormat(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            switch (ext)
            {
                case ".xls": return ExcelOutputFormat.Xls;
                case ".xlsm": return ExcelOutputFormat.Xlsm;
                default: return ExcelOutputFormat.Xlsx;
            }
        }

        private void MaterializeTempFromCurrentWorkbook()
        {
            // Новая книга сразу сохраняется во временный рабочий файл (чтобы сценарий с
            // WithHiddenWorkFile был одинаковым для шаблона и новой книги).
            string temp = CreateTempPath(nearTarget: _targetPath!, extension: Path.GetExtension(_targetPath!));
            try
            {
                Workbook.SaveAs(temp, GuessFormatByExtension(temp));
                Hide(temp);
                _tempFilePath = temp;
            }
            catch
            {
                TryDeleteTemp();
                throw;
            }
        }

        private static string CreateTempPath(string nearTarget, string extension)
        {
            string dir;
            try
            {
                dir = Directory.Exists(Path.GetDirectoryName(Path.GetFullPath(nearTarget)) ?? ".")
                    ? Path.GetDirectoryName(Path.GetFullPath(nearTarget))!
                    : Path.GetTempPath();
            }
            catch
            {
                dir = Path.GetTempPath();
            }

            if (string.IsNullOrEmpty(extension) || !extension.StartsWith("."))
                extension = ".xlsx";

            // имя вида ~EL~xxxxxxxx.xlsx (префикс «~» — традиционный признак временных файлов Office)
            string name = "~EL~" + Guid.NewGuid().ToString("N").Substring(0, 12) + extension;
            return Path.Combine(dir, name);
        }

        private static void Hide(string path)
        {
            try
            {
                var attrs = File.GetAttributes(path);
                File.SetAttributes(path, attrs | FileAttributes.Hidden);
            }
            catch
            {
                // атрибут Hidden недоступен на некоторых ФС — не критично
            }
        }

        private static string GuessStreamExtension(Stream stream)
        {
            // подпись ZIP (PK\x03\x04) → xlsx/xlsm; OLE CFBF (D0CF11E0) → xls
            try
            {
                if (!stream.CanSeek) return ".xlsx";
                long old = stream.Position;
                var buf = new byte[8];
                int read = stream.Read(buf, 0, 8);
                stream.Position = old;
                if (read >= 4 && buf[0] == 0x50 && buf[1] == 0x4B) return ".xlsx";
                if (read >= 8 && buf[0] == 0xD0 && buf[1] == 0xCF && buf[2] == 0x11 && buf[3] == 0xE0) return ".xls";
            }
            catch { /* best effort */ }
            return ".xlsx";
        }

        private void TryDeleteTemp()
        {
            if (_tempFilePath == null) return;
            try
            {
                if (File.Exists(_tempFilePath))
                {
                    File.SetAttributes(_tempFilePath, FileAttributes.Normal);
                    File.Delete(_tempFilePath);
                }
            }
            catch
            {
                // не переписываем исходную ошибку из-за неудалённого temp
            }
            _tempFilePath = null;
        }

        /// <summary>
        /// Идемпотентное освобождение: закрытие книги без сохранения; если финальное сохранение
        /// не выполнено — временный файл удаляется (критерий 10.1.3).
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try { Workbook?.Close(false); }
            catch { /* приложение могло быть уже освобождено */ }

            if (!_saved)
                TryDeleteTemp();
            else
                TryDeleteTemp(); // уже удалён в SaveTo — no-op
        }
    }
}
