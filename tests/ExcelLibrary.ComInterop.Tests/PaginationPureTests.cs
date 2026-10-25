using System;
using Xunit;

namespace ExcelLibrary.ComInterop.Tests
{
    /// <summary>
    /// Кроссплатформенные unit-тесты чистых (не COM) модулей этапов 3, 8, 9:
    /// нотация A1, параметры пагинации и pure-логика пагинатора.
    /// Закрывают технический долг, зафиксированный в Plan.md (примечания к этапам 5–9).
    /// </summary>
    public class PaginationPureTests
    {
        #region A1Address (Этап 3)

        [Theory]
        [InlineData("A1", 1, 1)]
        [InlineData("$B$12", 2, 12)]
        [InlineData("ab34", 28, 34)]
        [InlineData("S34", 19, 34)]
        [InlineData("AA100", 27, 100)]
        [InlineData("ZZ9", 676, 9)]
        public void A1_TryParse_ValidAddresses(string address, int expCol, int expRow)
        {
            Assert.True(A1Address.TryParse(address, out int col, out int row));
            Assert.Equal(expCol, col);
            Assert.Equal(expRow, row);
        }

        [Theory]
        [InlineData("")]
        [InlineData("  ")]
        [InlineData("1A")]
        [InlineData("A")]
        [InlineData("A0")]
        [InlineData("0A")]
        [InlineData("A-1")]
        [InlineData("Лист!A1")]
        public void A1_TryParse_InvalidAddresses_ReturnFalse(string address)
        {
            Assert.False(A1Address.TryParse(address, out _, out _));
        }

        [Fact]
        public void A1_Parse_ThrowsOnInvalid()
        {
            Assert.Throws<FormatException>(() => A1Address.Parse("not-a-cell"));
        }

        [Theory]
        [InlineData(1, 1, "A1")]
        [InlineData(19, 34, "S34")]
        [InlineData(27, 100, "AA100")]
        [InlineData(676, 9, "ZZ9")]
        public void A1_Format_RoundTrip(int col, int row, string expected)
        {
            Assert.Equal(expected, A1Address.Format(col, row));
            var (c, r) = A1Address.Parse(expected);
            Assert.Equal((col, row), (c, r));
        }

        [Fact]
        public void A1_Format_OutOfRange_Throws()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => A1Address.Format(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => A1Address.Format(1, 0));
        }

        [Theory]
        [InlineData("A", 1)]
        [InlineData("z", 26)]
        [InlineData("AA", 27)]
        [InlineData("AZ", 52)]
        [InlineData("BA", 53)]
        [InlineData("", 0)]
        [InlineData("1", 0)]
        public void A1_ColumnNameToIndex(string name, int expected)
        {
            Assert.Equal(expected, A1Address.ColumnNameToIndex(name));
        }

        [Fact]
        public void A1_ExtractSheetRefs_FromReferenceFormula()
        {
            // формулы вида из референсного шаблона: =Таблица!A11, ='Шаблон'!$B$12
            var formula = "=Таблица!A11 & Таблица!A12 & 'Другой'!$C$5";
            int count = 0;
            foreach (var _ in A1Address.ExtractSheetRefs(formula)) count++;
            Assert.Equal(3, count);
        }

        [Fact]
        public void A1_ExtractSheetRefs_EmptyOrNull_NoResults()
        {
            Assert.Empty(System.Linq.Enumerable.ToList(A1Address.ExtractSheetRefs(null!)));
            Assert.Empty(System.Linq.Enumerable.ToList(A1Address.ExtractSheetRefs("")));
        }

        [Theory]
        [InlineData("=Таблица!A11", "Таблица", true)]
        [InlineData("='Таблица'!$B$12", "Таблица", true)]
        [InlineData("=Таблица!A11", "ДругойЛист", false)]
        [InlineData("=СУММ(A1:A10)", "Таблица", false)]
        [InlineData("", "Таблица", false)]
        public void A1_ReferencesSheet(string formula, string sheet, bool expected)
        {
            Assert.Equal(expected, A1Address.ReferencesSheet(formula, sheet));
        }

        #endregion

        #region TemplateScanner helpers (Этап 8)

        [Theory]
        [InlineData("=((ROW(A1)-1-38)/35)+2", true)]   // формула номера листа из референса (S34)
        [InlineData("=((ROW(B5)-1)/31)+1", true)]
        [InlineData("=Таблица!A11", false)]            // ссылка на источник — не номер листа
        [InlineData("=ROW()", false)]                  // нет арифметики
        [InlineData("A1+1", false)]                    // не формула
        [InlineData("", false)]
        public void Scanner_LooksLikeSheetNumberFormula(string formula, bool expected)
        {
            Assert.Equal(expected, TemplateScanner.LooksLikeSheetNumberFormula(formula));
        }

        [Fact]
        public void Scanner_GetMatrixCell_SingleValueRange()
        {
            Assert.Equal("x", TemplateScanner.GetMatrixCell("x", 1, 1, 1, 1));
            Assert.Null(TemplateScanner.GetMatrixCell(null, 1, 1, 1, 1));
        }

        [Fact]
        public void Scanner_GetMatrixCell_TwoDimensionalArray()
        {
            var data = new object?[,] { { "a", "b" }, { "c", null } };
            Assert.Equal("a", TemplateScanner.GetMatrixCell(data, 2, 2, 1, 1));
            Assert.Equal("b", TemplateScanner.GetMatrixCell(data, 2, 2, 1, 2));
            Assert.Equal("c", TemplateScanner.GetMatrixCell(data, 2, 2, 2, 1));
            Assert.Null(TemplateScanner.GetMatrixCell(data, 2, 2, 2, 2));
            // выход за границы → null, без исключения (защита от «дырявых» массивов Interop)
            Assert.Null(TemplateScanner.GetMatrixCell(data, 3, 3, 3, 3));
        }

        #endregion

        #region PaginationOptions / Clone (Этапы 5, 8)

        [Fact]
        public void Options_Defaults_MatchReferenceTemplate()
        {
            var o = new PaginationOptions();
            Assert.Equal("Таблица", o.SourceSheetName);
            Assert.Equal("Для печати", o.FirstPageTemplateSheetName);
            Assert.Equal("Шаблон листа", o.SubsequentTemplateSheetName);
            // числовые параметры по умолчанию — null (автоопределение, п.5.6 запрет хардкода)
            Assert.Null(o.DataStartRowInSource);
            Assert.Null(o.RowsPerFirstPage);
            Assert.Null(o.RowsPerSubsequentPage);
            Assert.Null(o.SheetNumberCellAddress);
            Assert.True(o.HideServiceSheets);
            Assert.Equal(1000, o.MaxPages);
            Assert.Equal(200, o.WarnPagesThreshold);
        }

        [Fact]
        public void Options_Clone_IsDeepCopyOfArrays()
        {
            var o = new PaginationOptions { RowsPerFirstPage = 28 };
            var c = o.Clone();
            c.ServiceSheetsToHide[0] = "ИЗМЕНЕНО";
            c.RowsPerFirstPage = 99;

            Assert.Equal("Таблица", o.ServiceSheetsToHide[0]); // оригинал не затронут
            Assert.Equal(28, o.RowsPerFirstPage);
            Assert.Equal(99, c.RowsPerFirstPage);
        }

        [Fact]
        public void Builder_ChainsSettersIntoOptions()
        {
            // Build() — internal, доступен тестовой сборке через InternalsVisibleTo
            var b = new PaginationOptionsBuilder()
                .SourceSheet("Data")
                .FirstPageTemplate("P1")
                .SubsequentTemplate("Pn")
                .RowsPerFirstPage(10)
                .OutputSheetNamePattern("Стр {n}")
                .MaxPages(500);

            var o = b.Build();

            Assert.Equal("Data", o.SourceSheetName);
            Assert.Equal("P1", o.FirstPageTemplateSheetName);
            Assert.Equal("Pn", o.SubsequentTemplateSheetName);
            Assert.Equal(10, o.RowsPerFirstPage);
            Assert.Equal("Стр {n}", o.OutputSheetNamePattern);
            Assert.Equal(500, o.MaxPages);
        }

        #endregion

        #region SheetPaginator pure-логика (Этап 9)

        private static SheetPaginator CreatePaginator(PaginationOptions options)
        {
            // app не используется в проверяемых методах — допустим null
            var resolved = new ResolvedPaginationParameters
            {
                DataStartRow = options.DataStartRowInSource ?? 11,
                RowsPerFirstPage = options.RowsPerFirstPage ?? 11,
                RowsPerSubsequentPage = options.RowsPerSubsequentPage ?? 28,
            };
            return new SheetPaginator(null!, options, resolved);
        }

        [Fact]
        public void Paginator_ReferenceGeometry_PageOffsets()
        {
            // Геометрия референса 338535-00-90-000 ПЭ3: старт строки источника = 39,
            // первая страница = 11 строк, последующие = 28 (критерий 10.2.2: A39..A131 при T=100).
            var pag = CreatePaginator(new PaginationOptions
            {
                DataStartRowInSource = 39,
                RowsPerFirstPage = 11,
                RowsPerSubsequentPage = 28,
            });

            Assert.Equal(39, pag.FirstRowOfPage(1));   // лист 1: A39–A69
            Assert.Equal(50, pag.FirstRowOfPage(2));   // лист 2: A50–A77? — см. ниже
            Assert.Equal(78, pag.FirstRowOfPage(3));
            Assert.Equal(106, pag.FirstRowOfPage(4));
        }

        [Fact]
        public void Paginator_DefaultGeometry_OneHundredRows_FourPages()
        {
            // Дефолты автоопределения из референса: старт 11, первая 11, далее 28.
            // T=100: 11 + 28*3 = 95 < 100 ≤ 11 + 28*4 ⇒ 4 страницы (критерий 10.2.1).
            var pag = CreatePaginator(new PaginationOptions());
            Assert.Equal(11, pag.FirstRowOfPage(1));
            Assert.Equal(22, pag.FirstRowOfPage(2));
            Assert.Equal(50, pag.FirstRowOfPage(3));
            Assert.Equal(78, pag.FirstRowOfPage(4));

            int total = 100, covered = 11;
            int pages = 1;
            while (covered < total) { covered += 28; pages++; }
            Assert.Equal(4, pages);
        }

        [Fact]
        public void Paginator_FirstRowOfPage_MonotonicNoOverlap()
        {
            var pag = CreatePaginator(new PaginationOptions
            {
                DataStartRowInSource = 2,
                RowsPerFirstPage = 5,
                RowsPerSubsequentPage = 7,
            });
            // старты страниц: 2, 7, 14, 21 — строго возрастают, шаг = RowsPerSubsequentPage после 1-й
            Assert.Equal(2, pag.FirstRowOfPage(1));
            Assert.Equal(7, pag.FirstRowOfPage(2));
            Assert.Equal(14, pag.FirstRowOfPage(3));
            Assert.Equal(21, pag.FirstRowOfPage(4));
        }

        [Theory]
        [InlineData("Лист {n}", 3, "Лист 3")]
        [InlineData("Page-{n}", 12, "Page-12")]
        [InlineData("{n}", 7, "7")]
        public void Paginator_BuildSheetName_PatternReplacement(string pattern, int n, string expected)
        {
            var pag = CreatePaginator(new PaginationOptions { OutputSheetNamePattern = pattern });
            Assert.Equal(expected, pag.BuildSheetName(n));
        }

        [Fact]
        public void Paginator_BuildSheetName_TruncatesToExcelLimit()
        {
            var pag = CreatePaginator(new PaginationOptions
            {
                OutputSheetNamePattern = new string('X', 40) + "{n}"
            });
            Assert.Equal(31, pag.BuildSheetName(1).Length); // лимит имени листа в Excel — 31 символ
        }

        #endregion
    }
}
