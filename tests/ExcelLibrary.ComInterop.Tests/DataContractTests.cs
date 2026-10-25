using System;
using System.Collections.Generic;
using System.Text.Json;
using Xunit;

namespace ExcelLibrary.ComInterop.Tests
{
    /// <summary>
    /// Кроссплатформенные unit-тесты контрактов данных: JSON-модели схем (а)/(б) (Этап 3, п.4.1)
    /// и конструкторы/свойства иерархии исключений (Этап 1, п.1.7).
    /// </summary>
    public class DataContractTests
    {
        #region JSON-модели (схема (б): { "sheets": [...] })

        [Fact]
        public void JsonSchemaB_Deserialize_MapsAllFields()
        {
            string json = @"{
  ""sheets"": [
    {
      ""name"": ""Лист1"",
      ""startCell"": ""B2"",
      ""columns"": [""kd"", ""name""],
      ""data"": [ { ""kd"": 12, ""name"": ""Шифр"" }, { ""kd"": null, ""name"": ""Пусто"" } ]
    }
  ]
}";
            var doc = JsonSerializer.Deserialize<JsonSheetsDocument>(json);
            Assert.NotNull(doc);
            Assert.Single(doc!.Sheets);

            var sheet = doc.Sheets[0];
            Assert.Equal("Лист1", sheet.Name);
            Assert.Equal("B2", sheet.StartCell);
            Assert.Equal(new[] { "kd", "name" }, sheet.Columns);
            Assert.Equal(2, sheet.Data.Count);
            Assert.Equal("Шифр", sheet.Data[0]["name"]?.ToString());
            Assert.Null(sheet.Data[1]["kd"]);
        }

        [Fact]
        public void JsonSchemaB_Defaults_StartCellIsA1_SheetsIsEmpty()
        {
            var sheet = new JsonSheetData();
            Assert.Equal("A1", sheet.StartCell);
            Assert.Null(sheet.Name);
            Assert.Null(sheet.Columns);
            Assert.Empty(sheet.Data);

            var doc = new JsonSheetsDocument();
            Assert.Empty(doc.Sheets);
        }

        [Fact]
        public void JsonSchemaB_RoundTrip_PreservesStructure()
        {
            var doc = new JsonSheetsDocument
            {
                Sheets =
                {
                    new JsonSheetData
                    {
                        Name = "Данные",
                        StartCell = "A5",
                        Columns = new List<string> { "a", "b" },
                        Data = { new Dictionary<string, object?> { ["a"] = 1.5, ["b"] = "текст" } },
                    }
                }
            };

            string json = JsonSerializer.Serialize(doc);
            var parsed = JsonSerializer.Deserialize<JsonSheetsDocument>(json)!;
            Assert.Equal("Данные", parsed.Sheets[0].Name);
            Assert.Equal("A5", parsed.Sheets[0].StartCell);
            Assert.Equal(2, parsed.Sheets[0].Columns!.Count);
        }

        [Fact]
        public void JsonSchemaA_ArrayOfObjects_IsParseableAsRows()
        {
            // схема (а): плоский массив объектов → лист (парсится в Facade.FillFromJson напрямую)
            string json = @"[{""Version"":""1.2"",""KD"":10},{""Version"":""1.3"",""KD"":20}]";
            var rows = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json);
            Assert.Equal(2, rows!.Count);
            Assert.Equal("1.2", rows[0]["Version"].GetString());
            Assert.Equal(10, rows[0]["KD"].GetInt32());
        }

        #endregion

        #region Нормализация значений при записи (Этап 3, п.4.1/4.3)

        [Theory]
        [InlineData(null, null)]                    // null → пустая ячейка
        [InlineData("", "")]                        // пустая строка сохраняется как есть
        [InlineData("текст", "текст")]
        public void NormalizeValue_Strings(string? input, string? expected)
        {
            Assert.Equal(expected, ExcelLibraryFacade.NormalizeValue(input));
        }

        [Fact]
        public void NormalizeValue_Bool_MapsToExcelText()
        {
            Assert.Equal("True", ExcelLibraryFacade.NormalizeValue(true));
            Assert.Equal("False", ExcelLibraryFacade.NormalizeValue(false));
        }

        [Fact]
        public void NormalizeValue_Numbers_InvariantCulture()
        {
            // числа — инвариантная культура (без «,»/разделителей групп)
            Assert.Equal("42", ExcelLibraryFacade.NormalizeValue(42));
            Assert.Equal("1.5", ExcelLibraryFacade.NormalizeValue(1.5));
        }

        [Fact]
        public void NormalizeValue_DBNull_BecomesNull()
        {
            Assert.Null(ExcelLibraryFacade.NormalizeValue(DBNull.Value));
        }

        [Fact]
        public void NormalizeValue_DateTime_FormatsIsoLike()
        {
            var dt = new DateTime(2026, 10, 25, 13, 5, 9);
            Assert.Equal("2026-10-25 13:05:09", ExcelLibraryFacade.NormalizeValue(dt));
        }

        [Fact]
        public void NormalizeValue_Number_UsesInvariantCulture()
        {
            // даже в русскоязычной культуре десятичный разделитель — точка
            var saved = System.Globalization.CultureInfo.CurrentCulture;
            try
            {
                System.Globalization.CultureInfo.CurrentCulture =
                    new System.Globalization.CultureInfo("ru-RU");
                Assert.Equal("1234.5", ExcelLibraryFacade.NormalizeValue(1234.5));
            }
            finally
            {
                System.Globalization.CultureInfo.CurrentCulture = saved;
            }
        }

        #endregion

        #region Иерархия исключений (п.1.7) — свойства и сообщения

        [Fact]
        public void RangeOverflow_DefaultsAreZero_AndMessagePropagates()
        {
            var ex = new ExcelRangeOverflowException("переполнение");
            Assert.Equal(0, ex.RequestedCount);
            Assert.Equal(0, ex.MaximumCount);
            Assert.Equal("переполнение", ex.Message);
        }

        [Fact]
        public void RangeOverflow_WithInner_ExceptionChainIntact()
        {
            var inner = new InvalidOperationException("src");
            var ex = new ExcelRangeOverflowException("outer", inner, requestedCount: 131, maximumCount: 100);
            Assert.Same(inner, ex.InnerException);
            Assert.Equal(131, ex.RequestedCount);
            Assert.Equal(100, ex.MaximumCount);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("Для печати")]
        public void Template_CarriesElementName(string? elementName)
        {
            var ex = new ExcelTemplateException("нет элемента", elementName);
            Assert.Equal(elementName, ex.ElementName);
            Assert.IsAssignableFrom<ExcelLibraryException>(ex);
        }

        [Fact]
        public void BitnessMismatch_ExposesBothBitnesses()
        {
            var ex = new ExcelBitnessMismatchException(processBitness: 32, excelBitness: 64);
            Assert.Equal(32, ex.ProcessBitness);
            Assert.Equal(64, ex.ExcelBitness);

            var unknown = new ExcelBitnessMismatchException(processBitness: 64, excelBitness: null);
            Assert.Null(unknown.ExcelBitness);
        }

        [Fact]
        public void PaginationException_DerivesFromBase_AndSupportsInner()
        {
            var inner = new Exception("com");
            var ex = new ExcelPaginationException("откат выполнен", inner);
            Assert.Same(inner, ex.InnerException);
            Assert.IsAssignableFrom<ExcelLibraryException>(ex);
        }

        [Fact]
        public void AllExceptions_CatchableAsBaseType()
        {
            Exception[] all =
            {
                new ExcelRangeOverflowException("a", 1, 2),
                new ExcelTemplateException("b", "c"),
                new ExcelBitnessMismatchException(32, null),
                new ExcelPaginationException("d"),
            };
            foreach (var ex in all)
                Assert.IsAssignableFrom<ExcelLibraryException>(ex);
        }

        #endregion
    }
}
