using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Модуль автоматической пагинации печатных листов (Этап 9, п.5.5 ТЗ): разворачивает
    /// заполненный источник данных в N печатных листов на основе шаблонов первой и
    /// последующих страниц. Все параметры берутся из <see cref="ResolvedPaginationParameters"/>
    /// (автоопределение — <see cref="TemplateScanner"/>), хардкод значений референса запрещён (п.5.6).
    /// </summary>
    /// <remarks>
    /// Алгоритм:
    /// 1. Валидация листов/параметров → <see cref="ExcelPaginationException"/>.
    /// 2. T = число строк источника; распределение: T ≤ First ⇒ только первая страница,
    ///    иначе N = 1 + ceil((T − First) / Subsequent).
    /// 3. Первая страница остаётся как есть (PrintArea из шаблона не перезаписывается).
    /// 4. Копии через <c>Worksheet.Copy(After:)</c> (без Clipboard), перепись ссылок
    ///    <c>=Source!ref</c> пакетно (batch read формул → batch write), очистка «хвоста»
    ///    последней страницы, номер листа целым значением, перенос PageSetup без буфера обмена.
    /// 5. Откат при ошибке: все созданные копии удаляются (критерий 10.2.8).
    /// </remarks>
    public sealed class SheetPaginator
    {
        private readonly ComExcelApplication _app;
        private readonly PaginationOptions _options;
        private readonly ResolvedPaginationParameters _p;

        /// <summary>Число страниц, полученное при последнем вызове <see cref="Paginate"/>.</summary>
        public int LastPageCount { get; private set; }

        /// <summary>Предупреждения процесса пагинации (например, о большом числе страниц, п.5.7).</summary>
        public List<string> Warnings { get; } = new List<string>();

        /// <summary>Создаёт пагинатор для книги с готовыми параметрами.</summary>
        /// <param name="app">COM-приложение.</param>
        /// <param name="options">Опции пагинации.</param>
        /// <param name="resolved">Разрешённые (автоопределённые/override) параметры.</param>
        public SheetPaginator(ComExcelApplication app, PaginationOptions options, ResolvedPaginationParameters resolved)
        {
            _app = app;
            _options = options.Clone();
            _p = resolved;
        }

        /// <summary>
        /// Выполняет пагинацию книги: создаёт копии страниц после листа первой страницы.
        /// </summary>
        /// <param name="book">Рабочая книга (шаблон уже заполнен источником).</param>
        /// <returns>Число страниц N (включая первую).</returns>
        /// <exception cref="ExcelPaginationException">Валидация не пройдена или ошибка в процессе (выполнен откат копий).</exception>
        public int Paginate(ComWorkbook book)
        {
            Warnings.Clear();

            // —— 1. Валидация (п.5.5.1) ——
            var firstSheet = _app.GetWorksheet(book, _options.FirstPageTemplateSheetName);
            var sourceSheet = _app.GetWorksheet(book, _options.SourceSheetName);
            var subsTemplate = _app.GetWorksheet(book, _options.SubsequentTemplateSheetName);

            if (_p.RowsPerFirstPage <= 0 || _p.RowsPerSubsequentPage <= 0)
                throw new ExcelPaginationException("Параметры RowsPerFirstPage/RowsPerSubsequentPage должны быть > 0.");

            // начало данных не выше DataStartRowInSource (п.5.5.1, чек-лист 3.8)
            if (_p.DataStartRow < 1)
                throw new ExcelPaginationException($"DataStartRowInSource должен быть >= 1 (получено {_p.DataStartRow}).");

            // —— 2. Число строк источника T (п.5.5.2) ——
            int t = CountSourceRows(sourceSheet);

            // —— 3. Распределение (п.5.5.3, пограничные случаи п.5.7) ——
            int n = ComputePageCount(t, _p.RowsPerFirstPage, _p.RowsPerSubsequentPage);

            if (n == 1)
            {
                // только первая страница, копий нет (пустой источник — тоже)
                FinalizeFirstPage(firstSheet);
                ApplyServiceSheets(book);
                MaybeRecalculate(book);
                LastPageCount = 1;
                return 1;
            }

            int pagesSubsequent = n - 1;

            if (_options.WarnPagesThreshold > 0 && n >= _options.WarnPagesThreshold)
                Warnings.Add($"Пагинация создаст {n} страниц (порог предупреждения {_options.WarnPagesThreshold}); это может занять длительное время.");

            if (_options.MaxPages > 0 && n > _options.MaxPages)
                throw new ExcelPaginationException(
                    $"Требуется {n} страниц, что превышает лимит MaxPages={_options.MaxPages}. " +
                    "Увеличьте лимит или уменьшите объём данных (п.5.7 «большие объёмы»).");

            // —— 4..6. Создание копий с откатом при ошибке (п.5.7) ——
            var created = new List<ComWorksheet>();
            try
            {
                using (_app.BeginPerformanceScope()) // ScreenUpdating/Events/Calculation off (п.5.7)
                {
                    dynamic sheetsCol = ((dynamic)book.Raw).Worksheets;
                    object? lastAnchor = firstSheet.Raw; // копии идут сразу за первой страницей

                    for (int i = 1; i <= pagesSubsequent; i++)
                    {
                        int pageNumber = i + 1;
                        int page = pageNumber; // p в формулах плана

                        // 4.1 Copy(After:) целиком, без Clipboard (п.5.6)
                        ((dynamic)subsTemplate.Raw).Copy(After: lastAnchor);
                        int newIndex = ((int)((dynamic)lastAnchor).Index) + 1;
                        var copy = _app.Tracked(new ComWorksheet(_app, (object)sheetsCol[newIndex]));
                        created.Add(copy);
                        lastAnchor = copy.Raw;

                        // 4.2 имя по шаблону ({n} = номер страницы)
                        copy.Name = BuildSheetName(pageNumber);

                        // 4.3 перепись ссылок источника пакетно + очистка хвоста
                        RewriteSourceRefs(copy, page, t);

                        // 4.4 номер листа целым значением (дефолт по открытому вопросу 2)
                        WriteSheetNumber(copy, pageNumber);

                        // 4.5 перенос PageSetup с шаблона (сервис Этапа 5, без Clipboard)
                        PageSetupTransfer.Copy(subsTemplate, copy, transferPrintArea: false);

                        // 4.6 PrintArea копии — автоподбор по фактическим данным
                        AutoSetPrintArea(copy);
                    }
                }
            }
            catch (Exception ex)
            {
                // откат: удалить все созданные копии (временный файл удалит DocumentLifecycle)
                Rollback(created);
                if (ex is ExcelLibraryException) throw;
                throw new ExcelPaginationException(
                    "Пагинация прервана; все созданные копии листов удалены, шаблон не изменён.", ex);
            }

            FinalizeFirstPage(firstSheet);
            ApplyServiceSheets(book);
            MaybeRecalculate(book);
            LastPageCount = n;
            return n;
        }

        /// <summary>Число заполненных строк источника ниже DataStartRow (по первому непустому столбцу ссылок).</summary>
        private int CountSourceRows(ComWorksheet sourceSheet)
        {
            var used = sourceSheet.UsedRange;
            int lastUsedRow = used.Row + used.Rows - 1;
            if (lastUsedRow < _p.DataStartRow) return 0;

            // столбец для измерения: первый из столбцов ссылок источника (или A)
            int col = _p.SourceColumnsSubsequent.Length > 0 ? _p.SourceColumnsSubsequent[0]
                      : (_p.SourceColumnsFirstPage.Length > 0 ? _p.SourceColumnsFirstPage[0] : 1);

            string colName = A1Address.IndexToColumnName(col);
            var range = sourceSheet.Range($"{colName}{_p.DataStartRow}:{colName}{lastUsedRow}");
            object? data = ((dynamic)range.Raw).Value2;

            int rows = lastUsedRow - _p.DataStartRow + 1;
            int lastNonEmpty = 0;
            for (int r = 1; r <= rows; r++)
            {
                object? v = TemplateScanner.GetMatrixCell(data, rows, 1, r, 1);
                if (v != null && !(v is string sv && sv.Length == 0))
                    lastNonEmpty = r;
            }
            return lastNonEmpty;
        }

        /// <summary>
        /// Пакетная перепись ссылок <c>=Source!ref</c> на копии страницы (п.5.5.6.3, п.7):
        /// один batch-read формул UsedRange → трансформация → один batch-write.
        /// Ссылки на строки источника, которых нет (хвост последней страницы), очищаются (п.5.7).
        /// </summary>
        private void RewriteSourceRefs(ComWorksheet copy, int page, int totalSourceRows)
        {
            var used = copy.UsedRange;
            int rCount = used.Rows, cCount = used.Columns;
            if (rCount == 0 || cCount == 0) return;

            int baseRow = used.Row, baseCol = used.Column;
            object? formulas = ((dynamic)used.Raw).Formula;
            if (formulas == null) return;

            int firstRowOfPage = FirstRowOfPage(page);
            int minOriginal = _p.MinOriginalRefRowSubsequent;
            string source = _options.SourceSheetName;
            string escaped = Regex.Escape(source);
            // регэксп только для данного листа-источника (формулы на другие листы не трогаем, п.5.6)
            var refRegex = new Regex(
                @"(?<![\w!$:])('?)(" + escaped + @")\1!\$?([A-Za-z]{1,3})\$?(\d{1,7})",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            bool changed = false;
            var writes = new List<(int R, int C, string? Formula)>();

            for (int r = 1; r <= rCount; r++)
            {
                for (int c = 1; c <= cCount; c++)
                {
                    if (!(TemplateScanner.GetMatrixCell(formulas, rCount, cCount, r, c) is string f) || f.Length == 0)
                        continue;
                    if (!f.StartsWith("=")) continue;
                    if (!refRegex.IsMatch(f)) continue;

                    bool hasOverflow = false;
                    string nf = refRegex.Replace(f, m =>
                    {
                        int col = A1Address.ColumnNameToIndex(m.Groups[3].Value);
                        int origRow = int.Parse(m.Groups[4].Value);
                        int newRow = firstRowOfPage + (origRow - minOriginal);
                        if (newRow < _p.DataStartRow || newRow > _p.DataStartRow + totalSourceRows - 1)
                        {
                            hasOverflow = true; // ссылка на несуществующую строку источника
                        }
                        return $"{m.Groups[1].Value}{source}{m.Groups[1].Value}!" +
                               A1Address.Format(col, Math.Max(newRow, _p.DataStartRow));
                    });

                    writes.Add((r, c, hasOverflow ? (string?)null : nf));
                    changed = true;
                }
            }

            if (!changed) return;

            // сборка результирующей матрицы формул и ОДИН вызов записи Formula на UsedRange
            var result = new object?[rCount, cCount];
            for (int r = 1; r <= rCount; r++)
                for (int c = 1; c <= cCount; c++)
                    result[r - 1, c - 1] = TemplateScanner.GetMatrixCell(formulas, rCount, cCount, r, c);

            foreach (var (wr, wc, wf) in writes)
                result[wr - 1, wc - 1] = wf; // null ⇒ очистка ячейки (RemoveDuplicates: ClearContents у Range.Formula=null)

            ((dynamic)used.Raw).Formula = result;
            _ = baseRow; _ = baseCol; // базовые координаты не менялись (адреса относительные внутри массива)
        }

        /// <summary>
        /// Число страниц N по формуле п.5.5.3: T ≤ First ⇒ 1, иначе N = 1 + ceil((T − First) / Subsequent).
        /// Чистая функция без COM — используется <see cref="Paginate"/> и доступна для unit-тестов
        /// (в т.ч. критерий 10.2.1: T=100 при 28/31 ⇒ 4).
        /// </summary>
        /// <param name="totalRows">T — число заполненных строк источника.</param>
        /// <param name="rowsPerFirstPage">Строк на первой странице (должно быть &gt; 0).</param>
        /// <param name="rowsPerSubsequentPage">Строк на каждой последующей странице (должно быть &gt; 0).</param>
        /// <returns>Число страниц N ≥ 1.</returns>
        public static int ComputePageCount(int totalRows, int rowsPerFirstPage, int rowsPerSubsequentPage)
        {
            if (rowsPerFirstPage <= 0 || rowsPerSubsequentPage <= 0)
                throw new ExcelPaginationException("Параметры RowsPerFirstPage/RowsPerSubsequentPage должны быть > 0.");
            if (totalRows <= rowsPerFirstPage) return 1; // пустой источник или умещается на первой странице
            int remaining = totalRows - rowsPerFirstPage;
            return 1 + (remaining + rowsPerSubsequentPage - 1) / rowsPerSubsequentPage;
        }

        /// <summary>Номер строки источника, на которую указывает первая ссылка страницы p (п.5.5.6.3).</summary>
        internal int FirstRowOfPage(int page)
        {
            if (page <= 1) return _p.DataStartRow;
            return _p.DataStartRow + _p.RowsPerFirstPage + (page - 2) * _p.RowsPerSubsequentPage;
        }

        /// <summary>Проставляет номер листа целым значением (открытый вопрос 2: сломанная формула заменяется).</summary>
        private void WriteSheetNumber(ComWorksheet copy, int pageNumber)
        {
            if (string.IsNullOrEmpty(_p.SheetNumberCellAddress)) return;
            var cell = copy.Range(_p.SheetNumberCellAddress!);
            ((dynamic)cell.Raw).ClearContents(); // убираем старую формулу =((ROW(A1)-...)/...)+...
            ((dynamic)cell.Raw).Value2 = pageNumber;
        }

        /// <summary>Строит имя выходного листа по паттерну ({n} — номер страницы).</summary>
        internal string BuildSheetName(int pageNumber)
        {
            string pattern = string.IsNullOrEmpty(_options.OutputSheetNamePattern) ? "Лист {n}" : _options.OutputSheetNamePattern;
            string name = pattern.Replace("{n}", pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (name.Length > 31) name = name.Substring(0, 31); // лимит Excel
            return name;
        }

        /// <summary>Первая страница: изменения содержимого запрещены (п.5.5.5); PrintArea остаётся из шаблона.</summary>
        private void FinalizeFirstPage(ComWorksheet firstSheet)
        {
            // только применяем сохранённые настройки — никаких правок (чек-лист 9.6)
            // PrintArea первой страницы НЕ перезаписывается (п.5.2 чек-листа этапа 5)
            _ = firstSheet;
        }

        /// <summary>Автоподбор PrintArea по фактическим данным листа копии.</summary>
        private static void AutoSetPrintArea(ComWorksheet sheet)
        {
            var used = sheet.UsedRange;
            ((dynamic)sheet.Raw).PageSetup.PrintArea = used.Address;
        }

        /// <summary>Скрытие служебных листов после пагинации (п.5.5.7, критерий 10.2.6).</summary>
        private void ApplyServiceSheets(ComWorkbook book)
        {
            if (!_options.HideServiceSheets) return;
            foreach (var name in _options.ServiceSheetsToHide ?? Array.Empty<string>())
            {
                try
                {
                    var ws = _app.GetWorksheet(book, name);
                    ws.Visible = 0; // xlSheetHidden
                }
                catch (ExcelTemplateException)
                {
                    // лист отсутствует в книге — пропускаем (конфигурация может быть шире шаблона)
                }
            }
        }

        private void MaybeRecalculate(ComWorkbook book)
        {
            if (!_options.RecalculateOnFinish) return;
            try { ((dynamic)_app.Application!).CalculateFull(); }
            catch { /* пересчёт best-effort */ }
        }

        private void Rollback(List<ComWorksheet> created)
        {
            // удаляем копии в обратном порядке; ошибки игнорируем (книга может закрываться)
            for (int i = created.Count - 1; i >= 0; i--)
            {
                try { created[i].Delete(); }
                catch { }
            }
        }
    }
}
