using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Сканер шаблона (Этап 8, п.5.4): автоопределение <see cref="PaginationOptions"/> по формулам
    /// вида <c>=SourceSheet!A11</c> на листах-шаблонах — без хардкода значений референса (28/31/11/A2/A38).
    /// </summary>
    /// <remarks>
    /// Правила:
    /// - <b>DataStartRowInSource</b> = минимальный номер строки среди ссылок шаблона первой страницы;
    /// - <b>RowsPerFirstPage / RowsPerSubsequentPage</b> = количество <i>уникальных</i> строк источника,
    ///   на которые ссылается соответствующий шаблон (а не номер последней строки);
    /// - <b>SheetNumberCellAddress</b> — ячейка с маркером «Лист» в соседней ячейке либо с формулой
    ///   вида <c>=((ROW(...)-...)/...)+...</c>;
    /// - явные значения из опций имеют приоритет над автоопределением.
    /// </remarks>
    public static class TemplateScanner
    {
        /// <summary>
        /// Автоопределяет параметры пагинации по книге с шаблонами и применяет overrides из опций.
        /// </summary>
        /// <param name="app">COM-приложение.</param>
        /// <param name="book">Рабочая книга (открытый шаблон).</param>
        /// <param name="options">Опции пагинации (имена листов + возможные явные значения).</param>
        /// <returns>Разрешённые параметры для <see cref="SheetPaginator"/>.</returns>
        /// <exception cref="ExcelPaginationException">Шаблоны не содержат распознаваемых ссылок источника.</exception>
        public static ResolvedPaginationParameters Resolve(
            ComExcelApplication app, ComWorkbook book, PaginationOptions options)
        {
            var firstSheet = app.GetWorksheet(book, options.FirstPageTemplateSheetName);
            var subsSheet = app.GetWorksheet(book, options.SubsequentTemplateSheetName);

            var firstRefs = CollectSourceRows(app, firstSheet, options.SourceSheetName);
            var subsRefs = CollectSourceRows(app, subsSheet, options.SourceSheetName);

            if (firstRefs.Rows.Count == 0 && subsRefs.Rows.Count == 0)
                throw new ExcelPaginationException(
                    $"Не удалось автоопределить параметры пагинации: на листах " +
                    $"'{options.FirstPageTemplateSheetName}' и '{options.SubsequentTemplateSheetName}' " +
                    $"нет формул со ссылками на источник '{options.SourceSheetName}'.");

            // DataStartRow: минимум среди ссылок первого шаблона (если первый пуст — из последующего)
            int autoDataStart = firstRefs.Rows.Count > 0
                ? firstRefs.Rows.Min()
                : subsRefs.Rows.Min();

            int autoFirst = firstRefs.Rows.Count;   // уникальные строки источника (п.5.4.2)
            int autoSubs = subsRefs.Rows.Count;     // уникальные строки источника (п.5.4.3)

            // Ячейка номера листа: сначала в первом шаблоне, затем в последующем
            string? autoNumberCell = FindSheetNumberCell(app, firstSheet) ?? FindSheetNumberCell(app, subsSheet);

            var resolved = new ResolvedPaginationParameters
            {
                DataStartRow = options.DataStartRowInSource ?? autoDataStart,
                RowsPerFirstPage = options.RowsPerFirstPage ?? autoFirst,
                RowsPerSubsequentPage = options.RowsPerSubsequentPage ?? autoSubs,
                SheetNumberCellAddress = options.SheetNumberCellAddress ?? autoNumberCell,
                MinOriginalRefRowFirstPage = firstRefs.Rows.Count > 0 ? firstRefs.Rows.Min() : autoDataStart,
                MinOriginalRefRowSubsequent = subsRefs.Rows.Count > 0 ? subsRefs.Rows.Min() : autoDataStart,
                SourceColumnsFirstPage = firstRefs.Columns.OrderBy(c => c).ToArray(),
                SourceColumnsSubsequent = subsRefs.Columns.OrderBy(c => c).ToArray(),
                LastRefRowFirstPage = firstRefs.Rows.Count > 0 ? firstRefs.Rows.Max() : 0,
                LastRefRowSubsequent = subsRefs.Rows.Count > 0 ? subsRefs.Rows.Max() : 0,
                DataStartRowIsAuto = options.DataStartRowInSource == null,
                RowsPerFirstPageIsAuto = options.RowsPerFirstPage == null,
                RowsPerSubsequentPageIsAuto = options.RowsPerSubsequentPage == null,
                SheetNumberCellIsAuto = options.SheetNumberCellAddress == null,
            };

            if (resolved.RowsPerFirstPage <= 0)
                throw new ExcelPaginationException("RowsPerFirstPage должен быть больше нуля (автоопределение дало 0 — проверьте ссылки шаблона первой страницы).");
            if (resolved.RowsPerSubsequentPage <= 0)
                throw new ExcelPaginationException("RowsPerSubsequentPage должен быть больше нуля (автоопределение дало 0 — проверьте ссылки шаблона последующих страниц).");

            return resolved;
        }

        /// <summary>Собирает уникальные строки/столбцы источника, на которые ссылается лист.</summary>
        private static (HashSet<int> Rows, HashSet<int> Columns) CollectSourceRows(
            ComExcelApplication app, ComWorksheet sheet, string sourceSheetName)
        {
            var rows = new HashSet<int>();
            var cols = new HashSet<int>();

            // пакетное чтение формул всего UsedRange одним вызовом (п.7 производительности)
            var used = sheet.UsedRange;
            int rCount = used.Rows;
            int cCount = used.Columns;
            if (rCount == 0 || cCount == 0)
                return (rows, cols);

            object? formulas = ((dynamic)used.Raw).Formula;
            int baseRow = used.Row;
            int baseCol = used.Column;

            for (int r = 1; r <= rCount; r++)
            {
                for (int c = 1; c <= cCount; c++)
                {
                    object? cell = GetMatrixCell(formulas, rCount, cCount, r, c);
                    if (!(cell is string f) || f.Length == 0) continue;
                    if (!A1Address.ReferencesSheet(f, sourceSheetName)) continue;

                    foreach (var (col, row) in A1Address.ExtractSheetRefs(f))
                    {
                        rows.Add(row);
                        cols.Add(col);
                    }
                }
            }
            // baseRow/baseCol не участвуют: считаются именно строки ИСТОЧНИКА из ссылок (п.5.4)
            _ = baseRow; _ = baseCol;
            return (rows, cols);
        }

        /// <summary>Ищет ячейку номера листа: маркер «Лист» в соседней ячейке или формула с ROW().</summary>
        private static string? FindSheetNumberCell(ComExcelApplication app, ComWorksheet sheet)
        {
            var used = sheet.UsedRange;
            int rCount = used.Rows, cCount = used.Columns;
            if (rCount == 0 || cCount == 0) return null;

            int baseRow = used.Row, baseCol = used.Column;
            object? values = ((dynamic)used.Raw).Value2;
            object? formulas = ((dynamic)used.Raw).Formula;

            // 1) значение/формула вида =((ROW(...)...)/...)+... — кандидат на номер листа
            for (int r = 1; r <= rCount; r++)
            {
                for (int c = 1; c <= cCount; c++)
                {
                    if (GetMatrixCell(formulas, rCount, cCount, r, c) is string f && LooksLikeSheetNumberFormula(f))
                    {
                        // проверяем маркер «Лист» в любой из 4 соседей (как в референсе: S34 ← T33)
                        if (HasNeighborMarker(values, rCount, cCount, r, c))
                            return A1Address.Format(baseCol + c - 1, baseRow + r - 1);
                    }
                }
            }

            // 2) без привязки к формуле: ячейка справа/слева/сверху/снизу от маркера «Лист», содержащая число или формулу
            for (int r = 1; r <= rCount; r++)
            {
                for (int c = 1; c <= cCount; c++)
                {
                    if (!(GetMatrixCell(values, rCount, cCount, r, c) is string s) ||
                        !s.Trim().Equals("Лист", StringComparison.OrdinalIgnoreCase))
                        continue;

                    foreach (var (nr, nc) in Neighbors(r, c, rCount, cCount))
                    {
                        object? nf = GetMatrixCell(formulas, rCount, cCount, nr, nc);
                        object? nv = GetMatrixCell(values, rCount, cCount, nr, nc);
                        bool hasContent = (nf is string nfs && nfs.Length > 0) || nv != null;
                        if (hasContent)
                            return A1Address.Format(baseCol + nc - 1, baseRow + nr - 1);
                    }
                }
            }
            return null;
        }

        /// <summary>Формула номера листа: содержит ROW( и арифметику, но не ссылки на другие листы.</summary>
        internal static bool LooksLikeSheetNumberFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') return false;
            if (formula.IndexOf("ROW(", StringComparison.OrdinalIgnoreCase) < 0) return false;
            // вида =((ROW(A1)-1-38)/35)+2: деление и сложение с константами
            return formula.IndexOf('/') >= 0 && formula.IndexOf('+') >= 0 && !formula.Contains('!');
        }

        private static bool HasNeighborMarker(object? values, int rCount, int cCount, int r, int c)
        {
            foreach (var (nr, nc) in Neighbors(r, c, rCount, cCount))
            {
                if (GetMatrixCell(values, rCount, cCount, nr, nc) is string s &&
                    s.Trim().Equals("Лист", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private static IEnumerable<(int Row, int Col)> Neighbors(int r, int c, int rCount, int cCount)
        {
            if (r > 1) yield return (r - 1, c);
            if (r < rCount) yield return (r + 1, c);
            if (c > 1) yield return (r, c - 1);
            if (c < cCount) yield return (r, c + 1);
        }

        /// <summary>Доступ к элементу результата Value2/Formula (1-based массив или одиночное значение).</summary>
        internal static object? GetMatrixCell(object? data, int rows, int cols, int r, int c)
        {
            if (data == null) return null;
            if (rows == 1 && cols == 1) return data;
            if (data is Array arr)
            {
                if (arr.Rank != 2) return null;
                // Поддержка обоих форматов: 0-based object[,] (.NET Interop)
                // и 1-based SAFEARRAY double(1..n,) (VBA/Range.Value2).
                int i = r - 1 + arr.GetLowerBound(0);
                int j = c - 1 + arr.GetLowerBound(1);
                try { return arr.GetValue(i, j); }
                catch (IndexOutOfRangeException) { return null; }
            }
            return null;
        }
    }
}
