using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Реализация фасада <see cref="IExcelLibrary"/> (Этапы 2–9): делегирует COM-слой
    /// <see cref="ComExcelApplication"/>, жизненный цикл — <see cref="DocumentLifecycle"/>,
    /// пагинацию — <see cref="SheetPaginator"/>. Dispose идемпотентен (п.2.6).
    /// Порядок метаданных при пагинации (п.5.8): Version/KDNumber/Author/CreationDate — до,
    /// SheetsNumber — после (см. <see cref="Paginate(PaginationOptions)"/>).
    /// </summary>
    internal sealed class ExcelLibraryFacade : IExcelLibrary
    {
        private readonly DocumentLifecycle _doc;
        private readonly PaginationOptions? _paginationOptions;
        private readonly ComExcelApplication _app;
        private IDisposable? _appOwnership; // владеет COM-приложением (цепочка Dispose)
        private bool _disposed;
        private bool _saved;

        // лимиты листов Excel для проверки переполнения (п.4.3)
        private const int MaxRowsXlsx = 1_048_576;
        private const int MaxRowsXls = 65_536;
        private const int MaxCols = 16_384;

        internal ExcelLibraryFacade(DocumentLifecycle doc, PaginationOptions? paginationOptions)
        {
            _doc = doc;
            _app = FacadeOwnership.ExtractApp(doc);
            _paginationOptions = paginationOptions;
        }

        /// <inheritdoc />
        public bool IsSaved => _saved;

        private void EnsureAlive()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(IExcelLibrary), "Экземпляр библиотеки уже освобождён.");
        }

        private ComWorksheet Sheet(string name)
        {
            EnsureAlive();
            return _app.GetWorksheet(_doc.Workbook, name);
        }

        // ============================ Этап 3: заполнение данными ============================

        /// <inheritdoc />
        public void FillFromDataTable(string sheetName, string startCell, DataTable table, bool writeHeaders = true)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            int extraHeaderRows = writeHeaders ? 1 : 0;
            var matrix = ToMatrix(table, writeHeaders);
            WriteBlock(sheetName, startCell, matrix, table.Rows.Count + extraHeaderRows);
        }

        /// <inheritdoc />
        public void FillFromArray(string sheetName, string startCell, object[,] data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            var matrix = new object?[data.GetLength(0), data.GetLength(1)];
            for (int r = 0; r < data.GetLength(0); r++)
                for (int c = 0; c < data.GetLength(1); c++)
                    matrix[r, c] = NormalizeValue(data[r, c]);
            WriteBlock(sheetName, startCell, matrix, data.GetLength(0));
        }

        /// <inheritdoc />
        public void FillFromArray<T>(string sheetName, string startCell, IEnumerable<T> items, bool writeHeaders = false)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => p.CanRead).ToArray();
            var list = items.ToList();

            int rows = list.Count + (writeHeaders ? 1 : 0);
            var matrix = new object?[Math.Max(rows, 1), Math.Max(props.Length, 1)];

            int r0 = 0;
            if (writeHeaders)
            {
                for (int c = 0; c < props.Length; c++) matrix[0, c] = props[c].Name;
                r0 = 1;
            }
            for (int i = 0; i < list.Count; i++)
                for (int c = 0; c < props.Length; c++)
                    matrix[r0 + i, c] = NormalizeValue(props[c].GetValue(list[i]));

            WriteBlock(sheetName, startCell, matrix, rows);
        }

        /// <inheritdoc />
        public void FillFromJson(string sheetName, string json)
        {
            if (string.IsNullOrWhiteSpace(json)) throw new ArgumentException("JSON пуст.", nameof(json));
            JsonDocument docJson;
            try
            {
                docJson = JsonDocument.Parse(json);
            }
            catch (JsonException ex)
            {
                throw new ExcelLibraryException("Некорректный JSON при заполнении листа.", ex);
            }

            using (docJson)
            {
                var root = docJson.RootElement;

                // схема (б): { "sheets": [...] }
                if (root.ValueKind == JsonValueKind.Object &&
                    root.TryGetProperty("sheets", out var sheetsEl) &&
                    sheetsEl.ValueKind == JsonValueKind.Array)
                {
                    var model = JsonSerializer.Deserialize<JsonSheetsDocument>(json, JsonOpts);
                    if (model?.Sheets == null || model.Sheets.Count == 0)
                        throw new ExcelLibraryException("JSON схема (б): массив 'sheets' пуст или не читается.");
                    foreach (var sheetDto in model.Sheets)
                        ApplyJsonSheet(sheetDto.Name ?? sheetName, sheetDto.StartCell ?? "A1", sheetDto);
                    return;
                }

                // схема (а): массив объектов → один лист
                if (root.ValueKind == JsonValueKind.Array)
                {
                    var rowsDto = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json, JsonOpts);
                    var dto = new JsonSheetData { StartCell = "A1" };
                    var columns = new List<string>();
                    foreach (var row in rowsDto ?? new List<Dictionary<string, JsonElement>>())
                        foreach (var key in row.Keys)
                            if (!columns.Contains(key)) columns.Add(key);
                    dto.Columns = columns;
                    foreach (var row in rowsDto ?? new List<Dictionary<string, JsonElement>>())
                    {
                        var record = new Dictionary<string, object?>();
                        foreach (var kv in row) record[kv.Key] = JsonElementToObject(kv.Value);
                        dto.Data.Add(record);
                    }
                    ApplyJsonSheet(sheetName, dto.StartCell ?? "A1", dto);
                    return;
                }

                throw new ExcelLibraryException(
                    "JSON не соответствует поддерживаемым схемам: ожидается массив объектов (схема а) " +
                    "или объект { \"sheets\": [...] } (схема б).");
            }
        }

        private void ApplyJsonSheet(string sheetName, string startCell, JsonSheetData dto)
        {
            var columns = dto.Columns != null && dto.Columns.Count > 0
                ? dto.Columns
                : (dto.Data.Count > 0 ? new List<string>(dto.Data[0].Keys) : new List<string>());

            if (columns.Count == 0 || dto.Data.Count == 0) return;

            var matrix = new object?[dto.Data.Count, columns.Count];
            for (int r = 0; r < dto.Data.Count; r++)
                for (int c = 0; c < columns.Count; c++)
                    matrix[r, c] = NormalizeValue(dto.Data[r].ContainsKey(columns[c]) ? dto.Data[r][columns[c]] : null);

            WriteBlock(sheetName, startCell, matrix, dto.Data.Count);
        }

        private static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
        };

        private static object? JsonElementToObject(JsonElement el)
        {
            switch (el.ValueKind)
            {
                case JsonValueKind.String: return el.GetString();
                case JsonValueKind.Number:
                    if (el.TryGetInt32(out int i)) return i;
                    if (el.TryGetInt64(out long l)) return l;
                    return el.GetDouble();
                case JsonValueKind.True: return true;
                case JsonValueKind.False: return false;
                case JsonValueKind.Null:
                case JsonValueKind.Undefined: return null;
                default: return el.GetRawText();
            }
        }

        private static object?[,] ToMatrix(DataTable table, bool writeHeaders)
        {
            int cols = table.Columns.Count;
            int rows = table.Rows.Count + (writeHeaders ? 1 : 0);
            var matrix = new object?[Math.Max(rows, 1), Math.Max(cols, 1)];
            int r0 = 0;
            if (writeHeaders)
            {
                for (int c = 0; c < cols; c++) matrix[0, c] = table.Columns[c].ColumnName;
                r0 = 1;
            }
            for (int i = 0; i < table.Rows.Count; i++)
                for (int c = 0; c < cols; c++)
                    matrix[r0 + i, c] = NormalizeValue(table.Rows[i][c]);
            return matrix;
        }

        /// <summary>DBNull/null → null (пустая ячейка); всё остальное приводится к строке (текст, п.4.1).</summary>
        private static object? NormalizeValue(object? value)
        {
            if (value == null || value == DBNull.Value) return null;
            if (value is string s) return s;
            if (value is bool b) return b ? "True" : "False";
            if (value is DateTime dt) return dt.ToString("yyyy-MM-dd HH:mm:ss");
            if (value is IFormattable f) return f.ToString(null, System.Globalization.CultureInfo.InvariantCulture);
            return value.ToString();
        }

        /// <summary>
        /// Единая точка записи прямоугольника данных: валидация адреса, лимитов строк (.xls 65 536),
        /// запись одним вызовом Range.Value2 (п.4.3, п.7).
        /// </summary>
        private void WriteBlock(string sheetName, string startCell, object?[,] matrix, int dataRowCount)
        {
            var (startCol, startRow) = A1Address.Parse(startCell); // FormatException при некорректном адресе
            int rowCount = matrix.GetLength(0);
            int colCount = matrix.GetLength(1);
            if (rowCount == 0 || colCount == 0) return;

            int maxRows = _doc.IsXlsTarget ? MaxRowsXls : MaxRowsXlsx;
            long lastRow = (long)startRow + rowCount - 1;
            long lastCol = (long)startCol + colCount - 1;
            if (lastRow > maxRows)
                throw new ExcelRangeOverflowException(
                    $"Запись на '{sheetName}' от ячейки {startCell} выходит за лимит {maxRows} строк" +
                    (_doc.IsXlsTarget ? " формата .xls (п.4.3)." : "."),
                    (int)Math.Min(lastRow, int.MaxValue), maxRows);
            if (lastCol > MaxCols)
                throw new ExcelRangeOverflowException(
                    $"Запись на '{sheetName}' выходит за лимит {MaxCols} столбцов.", (int)lastCol, MaxCols);

            // проверка лимита и для открытой книги .xls (дефолт открытого вопроса 4)
            _doc.EnsureWithinSheetLimit(startRow, rowCount);

            var sheet = Sheet(sheetName);
            string address = $"{A1Address.Format(startCol, startRow)}:{A1Address.Format(lastCol >= 1 ? (int)lastCol : startCol, (int)Math.Min(lastRow, int.MaxValue))}";
            var range = sheet.Range(address);

            var payload = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
                for (int c = 0; c < colCount; c++)
                    payload[r, c] = matrix[r, c]!;

            // одна пакетная запись Value2 (п.7)
            ((dynamic)range.Raw).Value2 = payload;

            // все значения пишутся как текст (п.4.1 — типизация на следующем этапе, п.8 ТЗ)
            ((dynamic)range.Raw).NumberFormat = "@";
        }

        // ============================ Этап 4: именованные диапазоны ============================

        /// <inheritdoc />
        public void SetNamedValue(string key, object value)
        {
            if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("Имя диапазона не задано.", nameof(key));
            EnsureAlive();

            dynamic names = ((dynamic)_doc.Workbook.Raw).Names;

            // значение пишется как текст в первую ячейку листа (формулы не поддерживаются, п.8)
            var targetSheet = ResolveNamedTargetSheet();
            ComRange cell = targetSheet.Cell("A1");
            ((dynamic)cell.Raw).Value2 = value?.ToString();

            // перезапись существующего имени без ошибки (критерий 10.1.4)
            try { names.Item(key).Delete(); }
            catch (COMException) { /* имя отсутствовало */ }
            catch (Exception ex) when (ex.GetType().Name == "CommonException") { /* имя отсутствовало */ }

            string refersTo = $"'{targetSheet.Name}'!$A$1";
            names.Add(key, Type.Missing, Type.Missing, Type.Missing, refersTo);
        }

        /// <inheritdoc />
        public void SetNamedValues(IDictionary<string, object> values)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            foreach (var kv in values)
                SetNamedValue(kv.Key, kv.Value);
        }

        /// <summary>Лист для значений имён: служебный «Info» (референс) или первый лист книги.</summary>
        private ComWorksheet ResolveNamedTargetSheet()
        {
            try { return Sheet("Info"); }
            catch (ExcelTemplateException)
            {
                dynamic sheets = ((dynamic)_doc.Workbook.Raw).Worksheets;
                return _app.Tracked(new ComWorksheet(_app, (object)sheets[1]));
            }
        }

        /// <summary>
        /// Читает значение глобального имени; бросает исключение для многоячеечных диапазонов (п.4 чек-лист 4.4).
        /// </summary>
        /// <param name="key">Имя диапазона.</param>
        /// <returns>Значение ячейки.</returns>
        public object? GetNamedValue(string key)
        {
            EnsureAlive();
            var range = _app.GetNamedRange(_doc.Workbook, key);
            int cells = range.Rows * range.Columns;
            if (cells != 1)
                throw new ExcelTemplateException(
                    $"Именованный диапазон '{key}' занимает {cells} ячеек; библиотека поддерживает только одноячеечные имена (п.4.2).", key);
            return ((dynamic)range.Raw).Value2;
        }

        // ============================ Этап 6: изображения ============================

        /// <inheritdoc />
        public void InsertImage(string sheetName, string cell, object source,
            double? widthPx = null, double? heightPx = null,
            ImageSizeMode sizeMode = ImageSizeMode.AbsolutePoints,
            int widthCells = 1, int heightCells = 1)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            EnsureAlive();

            string filePath = MaterializeImageSource(source);

            var ws = Sheet(sheetName);
            var anchor = ws.Range(cell);
            dynamic shapes = ((dynamic)ws.Raw).Shapes;

            // LinkToFile=false, SaveWithDocument=true — изображение встраивается (критерий 10.1.5)
            dynamic picture = shapes.AddPicture(filePath, LinkToFile: false, SaveWithDocument: true,
                Left: (double)((dynamic)anchor.Raw).Left, Top: (double)((dynamic)anchor.Raw).Top);

            if (sizeMode == ImageSizeMode.AbsolutePoints)
            {
                // px→pt: 1 px = 0.75 pt при 96 dpi
                if (widthPx.HasValue) picture.Width = widthPx.Value * 0.75;
                if (heightPx.HasValue) picture.Height = heightPx.Value * 0.75;
            }
            else
            {
                var (col, row) = A1Address.Parse(cell);
                dynamic fromR = ((dynamic)ws.Raw).Range[cell];
                dynamic toR = ((dynamic)ws.Raw).Range[
                    $"{A1Address.Format(col + widthCells - 1, row + heightCells - 1)}"];
                double left = (double)fromR.Left;
                double top = (double)fromR.Top;
                double right = (double)toR.Left + (double)toR.Width;
                double bottom = (double)toR.Top + (double)toR.Height;
                picture.Left = left;
                picture.Top = top;
                picture.Width = right - left;
                picture.Height = bottom - top;
            }

            picture.Placement = 1; // xlMoveAndSize (п.4.5)

            // временный файл изображения удаляем — картинка уже встроена в документ
            if (_tempImages.TryGetValue(filePath, out bool isTemp) && isTemp)
            {
                try { File.Delete(filePath); } catch { }
                _tempImages.Remove(filePath);
            }
        }

        private readonly Dictionary<string, bool> _tempImages = new Dictionary<string, bool>();

        /// <summary>Материализует источник изображения в путь (path/Stream/byte[]), проверяет формат png/jpg/bmp.</summary>
        private string MaterializeImageSource(object source)
        {
            string? path = null;
            bool isTemp = false;

            switch (source)
            {
                case string s:
                    if (!File.Exists(s)) throw new FileNotFoundException("Файл изображения не найден.", s);
                    path = Path.GetFullPath(s);
                    break;
                case byte[] bytes:
                    path = WriteTempImage(bytes, ".png");
                    isTemp = true;
                    break;
                case Stream stream:
                    using (var ms = new MemoryStream())
                    {
                        stream.CopyTo(ms);
                        path = WriteTempImage(ms.ToArray(), ".png");
                    }
                    isTemp = true;
                    break;
                default:
                    throw new ArgumentException("Источник изображения должен быть string (путь), Stream или byte[].");
            }

            string ext = Path.GetExtension(path!).ToLowerInvariant();
            if (ext != ".png" && ext != ".jpg" && ext != ".jpeg" && ext != ".bmp")
            {
                if (isTemp) { try { File.Delete(path!); } catch { } }
                throw new ExcelLibraryException($"Неподдерживаемый формат изображения '{ext}'. Допустимы .png, .jpg, .bmp (п.4.5).");
            }

            _tempImages[path!] = isTemp;
            return path!;
        }

        private static string WriteTempImage(byte[] bytes, string ext)
        {
            string tmp = Path.Combine(Path.GetTempPath(), "~EL~img" + Guid.NewGuid().ToString("N") + ext);
            File.WriteAllBytes(tmp, bytes);
            return tmp;
        }

        // ============================ Этап 5: печать, колонтитулы, листы ============================

        /// <inheritdoc />
        public void SetPrintArea(string sheetName, bool auto = true, string? printArea = null)
        {
            var ws = Sheet(sheetName);
            if (auto)
            {
                ((dynamic)ws.Raw).PageSetup.PrintArea = ws.UsedRange.Address;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(printArea))
                    throw new ArgumentException("printArea обязателен при auto=false.");
                ((dynamic)ws.Raw).PageSetup.PrintArea = FixRelativeAddress(ws, printArea);
            }
        }

        /// <summary>Приводит относительный адрес («A1:D10») к абсолютному с именем листа — требование PrintTitleRows/Columns.</summary>
        private static string FixRelativeAddress(ComWorksheet ws, string a1)
        {
            if (a1.StartsWith("$") || a1.Contains("!")) return a1;
            var parts = a1.Split(':');
            string Abs(string p)
            {
                var (c, r) = A1Address.Parse(p);
                return "$" + A1Address.IndexToColumnName(c) + "$" + r;
            }
            string body = parts.Length == 1 ? Abs(parts[0]) : Abs(parts[0]) + ":" + Abs(parts[1]);
            return $"'{ws.Name}'!{body}";
        }

        /// <inheritdoc />
        public void SetHeaderFooter(string sheetName, HeaderFooter? header = null, HeaderFooter? footer = null)
        {
            var ws = Sheet(sheetName);
            dynamic ps = ((dynamic)ws.Raw).PageSetup;
            if (header != null)
            {
                if (header.Left != null) ps.LeftHeader = header.Left;
                if (header.Center != null) ps.CenterHeader = header.Center;
                if (header.Right != null) ps.RightHeader = header.Right;
            }
            if (footer != null)
            {
                if (footer.Left != null) ps.LeftFooter = footer.Left;
                if (footer.Center != null) ps.CenterFooter = footer.Center;
                if (footer.Right != null) ps.RightFooter = footer.Right;
            }
        }

        /// <inheritdoc />
        public (HeaderFooter Header, HeaderFooter Footer) GetHeaderFooter(string sheetName)
        {
            var ws = Sheet(sheetName);
            dynamic ps = ((dynamic)ws.Raw).PageSetup;
            return (
                new HeaderFooter { Left = (string?)ps.LeftHeader, Center = (string?)ps.CenterHeader, Right = (string?)ps.RightHeader },
                new HeaderFooter { Left = (string?)ps.LeftFooter, Center = (string?)ps.CenterFooter, Right = (string?)ps.RightFooter });
        }

        /// <inheritdoc />
        public void SetPrintTitles(string sheetName, string? rows, string? columns)
        {
            var ws = Sheet(sheetName);
            dynamic ps = ((dynamic)ws.Raw).PageSetup;
            if (rows != null) ps.PrintTitleRows = FixRelativeAddress(ws, rows.Contains(":") ? rows : rows + ":" + rows);
            if (columns != null) ps.PrintTitleColumns = FixRelativeAddress(ws, columns.Contains(":") ? columns : columns + ":" + columns);
        }

        /// <inheritdoc />
        public void AddSheet(string name, string? after = null, string? modelSheet = null)
        {
            EnsureAlive();
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Имя листа не задано.", nameof(name));

            // проверка уникальности имени
            foreach (var existing in _app.GetWorksheets(_doc.Workbook))
                if (string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase))
                    throw new ExcelTemplateException($"Лист '{name}' уже существует в книге.", name);

            dynamic sheets = ((dynamic)_doc.Workbook.Raw).Worksheets;
            ComWorksheet? afterWs = after != null ? Sheet(after) : null;
            dynamic newWs = afterWs != null
                ? sheets.Add(After: afterWs.Raw)
                : sheets.Add(Type.Missing, Type.Missing, Type.Missing, 3 /* xlWorkSheet */);
            var created = _app.Tracked(new ComWorksheet(_app, (object)newWs));
            created.Name = name;

            if (modelSheet != null)
            {
                var model = Sheet(modelSheet);
                PageSetupTransfer.Copy(model, created, transferPrintArea: true); // без Clipboard (п.4.6)
            }
        }

        /// <inheritdoc />
        public void RemoveSheet(string name)
        {
            EnsureAlive();
            var ws = Sheet(name);

            dynamic sheets = ((dynamic)_doc.Workbook.Raw).Worksheets;
            int count = (int)sheets.Count;
            if (count <= 1)
                throw new ExcelTemplateException("Нельзя удалить единственный лист книги.", name);

            // защита листа (CheckProtection через Sheet.Protect: AttemptPassword-проверка недоступна —
            // используем свойство Protection via COM: ProtectContents через reflection best-effort)
            try
            {
                dynamic protection = ((dynamic)ws.Raw).Protection;
                bool protectedSheet = (bool)protection.ProtectContents;
                if (protectedSheet)
                    throw new ExcelTemplateException($"Лист '{name}' защищён — удаление запрещено (п.4.8).", name);
            }
            catch (COMException) { /* определение защиты best-effort */ }

            ws.Delete();
        }

        // ============================ Этапы 8–9: пагинация ============================

        /// <inheritdoc />
        public int Paginate()
        {
            if (_paginationOptions == null)
                throw new InvalidOperationException(
                    "Пагинация не настроена: вызовите WithPagination(...) в билдере или Paginate(options).");
            return Paginate(_paginationOptions);
        }

        /// <inheritdoc />
        public int Paginate(PaginationOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            EnsureAlive();

            // порядок метаданных (п.5.8): Version/KDNumber/Author/CreationDate — ДО пагинации
            // (их устанавливает клиент через SetNamedValue до этого вызова); SheetsNumber — ПОСЛЕ.
            var resolved = TemplateScanner.Resolve(_app, _doc.Workbook, options);
            var paginator = new SheetPaginator(_app, options, resolved);
            int pages = paginator.Paginate(_doc.Workbook);

            // SheetsNumber — после пагинации (дефолт открытого вопроса 3, критерий 10.2.7)
            try { SetNamedValue("SheetsNumber", pages); }
            catch (ExcelLibraryException) { /* если имя недоступно — не считаем пагинацию проваленной */ }

            LastWarnings = paginator.Warnings.ToList();
            return pages;
        }

        /// <summary>Предупреждения последней пагинации (п.5.7 большие объёмы).</summary>
        public List<string> LastWarnings { get; private set; } = new List<string>();

        // ============================ Этап 7: PDF ============================

        /// <inheritdoc />
        public void ExportSheetToPdf(string sheetName, string pdfPath, bool overwrite = true,
            PdfQuality quality = PdfQuality.Standard, bool includeMarkup = false)
        {
            EnsureAlive();
            if (string.IsNullOrWhiteSpace(pdfPath)) throw new ArgumentException("Путь PDF не задан.", nameof(pdfPath));
            if (!pdfPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Целевой файл PDF должен иметь расширение .pdf.");

            string full = Path.GetFullPath(pdfPath);
            if (File.Exists(full) && !overwrite)
                throw new ExcelLibraryException($"PDF '{full}' существует, а overwrite=false (п.4.7).");

            var ws = Sheet(sheetName);
            string? tmp = null;
            try
            {
                // экспорт во временный файл, затем перенос на целевой — при неудаче ничего не портим
                tmp = Path.Combine(Path.GetDirectoryName(full) ?? Path.GetTempPath(),
                    "~EL~pdf" + Guid.NewGuid().ToString("N") + ".pdf");

                dynamic fixedFormatQuality = (int)quality; // xlQualityMinimum=0 / xlQualityStandard=1; High → Standard+FromExt
                ws.Activate();
                ((dynamic)ws.Raw).ExportAsFixedFormat(
                    Type: 0,                       // xlTypePDF
                    Filename: tmp,
                    Quality: (int)quality == 0 ? 0 : 1, // xlQualityMinimum/xlQualityStandard
                    IncludeDocProperties: true,
                    IgnorePrintAreas: !includeMarkup ? false : false, // разметку учитываем всегда; флаг — про printing of markup comments
                    From: Type.Missing,
                    To: Type.Missing,
                    OpenAfterPublish: false,
                    FixedFormatExtClassPtr: Type.Missing);

                if (File.Exists(full) && overwrite)
                    File.Delete(full);
                File.Move(tmp, full);
                tmp = null;
            }
            catch (Exception ex) when (!(ex is ExcelLibraryException))
            {
                throw new ExcelLibraryException($"Не удалось экспортировать лист '{sheetName}' в PDF '{full}'.", ex);
            }
            finally
            {
                if (tmp != null) { try { File.Delete(tmp); } catch { } }
            }
        }

        // ============================ Сохранение / Dispose ============================

        /// <inheritdoc />
        public void Save()
        {
            EnsureAlive();
            _doc.SaveToTarget();
            _saved = true;
        }

        /// <inheritdoc />
        public void SaveTo(string path)
        {
            EnsureAlive();
            _doc.SaveTo(path);
            _saved = true;
        }

        /// <summary>
        /// Идемпотентное освобождение: закрывает книгу, удаляет несохранённый временный рабочий файл
        /// (критерий 10.1.3), затем освобождает COM-приложение (Quit для созданного нами EXCEL.EXE,
        /// чужие процессы не трогаются).
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _doc.Dispose();
            _appOwnership?.Dispose();
            _appOwnership = null;
        }
    }

    /// <summary>
    /// Служебный доступ к внутренним связям фасада (без публичного API): извлечение
    /// COM-приложения из жизненного цикла документа и прикрепление владельца приложения.
    /// </summary>
    internal static class FacadeOwnership
    {
        /// <summary>Достаёт ComExcelApplication из DocumentLifecycle (internal-поле _app).</summary>
        public static ComExcelApplication ExtractApp(DocumentLifecycle doc)
        {
            var field = typeof(DocumentLifecycle).GetField("_app", BindingFlags.NonPublic | BindingFlags.Instance);
            return (ComExcelApplication)field!.GetValue(doc)!;
        }

        /// <summary>Передаёт фасаду владение COM-приложением для цепочки Dispose.</summary>
        public static void Attach(ExcelLibraryFacade facade, IDisposable appOwnership)
        {
            var field = typeof(ExcelLibraryFacade).GetField("_appOwnership", BindingFlags.NonPublic | BindingFlags.Instance);
            field!.SetValue(facade, appOwnership);
        }
    }

    /// <summary>
    /// Сервис переноса настроек печати между листами (Этап 5, п.4.6): только свойства PageSetup,
    /// без Clipboard/PasteSpecial (п.5.6). Используется пагинатором и AddSheet.
    /// </summary>
    public static class PageSetupTransfer
    {
        private static readonly string[] Properties =
        {
            "Orientation", "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
            "Zoom", "FitToPagesWide", "FitToPagesTall", "CenterHorizontally", "CenterVertically",
            "HeaderMargin", "FooterMargin", "PaperSize", "FirstPageNumber",
            "LeftHeader", "CenterHeader", "RightHeader", "LeftFooter", "CenterFooter", "RightFooter",
            "PrintTitleRows", "PrintTitleColumns",
        };

        /// <summary>Перенести настройки печати с листа-источника на лист-приёмник.</summary>
        /// <param name="source">Лист-образец.</param>
        /// <param name="target">Целевой лист.</param>
        /// <param name="transferPrintArea">Переносить ли PrintArea (пагинация использует автоподбор — false).</param>
        /// <exception cref="ExcelTemplateException">Свойство PageSetup недоступно.</exception>
        public static void Copy(ComWorksheet source, ComWorksheet target, bool transferPrintArea = true)
        {
            dynamic srcPs = ((dynamic)source.Raw).PageSetup;
            dynamic dstPs = ((dynamic)target.Raw).PageSetup;

            foreach (var prop in Properties)
            {
                try
                {
                    object? value = ((object)srcPs).GetType().GetProperty(prop)?.GetValue(srcPs);
                    if (value == null) continue;
                    ((object)dstPs).GetType().GetProperty(prop)?.SetValue(dstPs, UnwrapMissing(value));
                }
                catch (COMException)
                {
                    // отдельные свойства могут быть недоступны в конкретной версии Excel — пропускаем
                }
                catch (MissingMemberException)
                {
                }
            }

            if (transferPrintArea)
            {
                try
                {
                    string? area = (string?)srcPs.PrintArea;
                    if (!string.IsNullOrEmpty(area))
                        dstPs.PrintArea = RebasePrintArea(area, source.Name, target.Name);
                }
                catch (COMException) { }
            }
        }

        private static object UnwrapMissing(object v) => v ?? Type.Missing;

        /// <summary>Подменяет имя листа в адресе PrintArea ('Старое'!$A$1:$D$10 → 'Новое'!...).</summary>
        internal static string RebasePrintArea(string area, string oldSheet, string newSheet)
        {
            if (area.Contains(oldSheet))
                return area.Replace("'" + oldSheet + "'", "'" + newSheet + "'")
                            .Replace(oldSheet + "!", newSheet + "!");
            return area;
        }
    }

    /// <summary>Модель JSON-схемы (б): корневой объект с массивом листов (п.4.1, описана в README).</summary>
    public sealed class JsonSheetsDocument
    {
        /// <summary>Листы для заполнения.</summary>
        [JsonPropertyName("sheets")]
        public List<JsonSheetData> Sheets { get; set; } = new List<JsonSheetData>();
    }

    /// <summary>Один лист в JSON-схеме (б): name/startCell/columns/data.</summary>
    public sealed class JsonSheetData
    {
        /// <summary>Имя листа назначения.</summary>
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        /// <summary>Верхняя левая ячейка записи (по умолчанию A1).</summary>
        [JsonPropertyName("startCell")]
        public string? StartCell { get; set; } = "A1";

        /// <summary>Порядок столбцов (ключи объектов из data).</summary>
        [JsonPropertyName("columns")]
        public List<string>? Columns { get; set; }

        /// <summary>Строки данных: объекты ключ→значение.</summary>
        [JsonPropertyName("data")]
        public List<Dictionary<string, object?>> Data { get; set; } = new List<Dictionary<string, object?>>();
    }
}
