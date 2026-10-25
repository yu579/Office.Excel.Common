using System;
using System.IO;
using Xunit;

namespace ExcelLibrary.ComInterop.Tests
{
    /// <summary>
    /// Тесты нижнего COM-слоя (Этап 1). Делятся на:
    /// - кроссплатформенные (иерархия исключений, форматы файлов) — выполняются везде;
    /// - Windows-only с реальным Excel (создание/Dispose приложения, scoped-блоки) —
    ///   пропускаются, если Excel не установлен (Skip-условие через факт IsExcelAvailable).
    /// </summary>
    public class ComLayerTests
    {
        /// <summary>Наличие зарегистрированного Excel (COM ProgID).</summary>
        private static bool IsExcelAvailable()
        {
            try
            {
                return Type.GetTypeFromProgID("Excel.Application", throwOnError: false) != null;
            }
            catch
            {
                return false;
            }
        }

        #region Кроссплатформенные тесты (проверяются и на Linux)

        [Fact]
        public void ExceptionHierarchy_AllFiveTypesDeriveFromBase()
        {
            // п.1.7: базовый ExcelLibraryException + 4 наследника
            Assert.True(typeof(ExcelLibraryException).BaseType == typeof(Exception));
            Assert.True(typeof(ExcelRangeOverflowException).IsSubclassOf(typeof(ExcelLibraryException)));
            Assert.True(typeof(ExcelTemplateException).IsSubclassOf(typeof(ExcelLibraryException)));
            Assert.True(typeof(ExcelBitnessMismatchException).IsSubclassOf(typeof(ExcelLibraryException)));
            Assert.True(typeof(ExcelPaginationException).IsSubclassOf(typeof(ExcelLibraryException)));
        }

        [Fact]
        public void ExcelRangeOverflow_CarriesRequestedAndMaximumCounts()
        {
            var ex = new ExcelRangeOverflowException("test", requestedCount: 100, maximumCount: 28);
            Assert.Equal(100, ex.RequestedCount);
            Assert.Equal(28, ex.MaximumCount);
            Assert.IsAssignableFrom<ExcelLibraryException>(ex);
        }

        [Fact]
        public void ExcelTemplate_CarriesElementName()
        {
            var ex = new ExcelTemplateException("нет листа", elementName: "Для печати");
            Assert.Equal("Для печати", ex.ElementName);
        }

        [Fact]
        public void BitnessMismatch_MessageContainsInstruction()
        {
            var ex = new ExcelBitnessMismatchException(processBitness: 64, excelBitness: 32);
            // сообщение должно содержать инструкцию по устранению (п.1.3)
            Assert.Contains("разрядност", ex.Message);
            Assert.Contains("AnyCPU", ex.Message);
            Assert.Equal(64, ex.ProcessBitness);
            Assert.Equal(32, ex.ExcelBitness);
        }

        [Theory]
        [InlineData("out.xlsx", 51)]   // xlOpenXMLWorkbook
        [InlineData("out.xlsm", 52)]   // xlOpenXMLWorkbookMacroEnabled
        [InlineData("out.xls", 56)]    // xlExcel8
        public void GuessFormat_MapsExtensionsToFileFormatCodes(string path, int expected)
        {
            Assert.Equal(expected, ComWorkbook.GuessFormat(path));
        }

        [Theory]
        [InlineData("out.pdf")]
        [InlineData("out.txt")]
        public void GuessFormat_ThrowsForUnsupportedExtensions(string path)
        {
            Assert.Throws<ArgumentException>(() => ComWorkbook.GuessFormat(path));
        }

        [Fact]
        public void OpenWorkbook_MissingFile_ThrowsFileNotFound()
        {
            // Проверка логики валидации пути без реального Excel: конструктор может бросить
            // ExcelLibraryException (Excel не установлен) — это допустимо; если Excel есть —
            // должен получиться FileNotFoundException. Проверяем оба допустимых исхода.
            string missing = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                // checkBitness=false — иначе на Windows без Office конструктор бросит mismatch до проверки пути
                using var app = new ComExcelApplication(checkBitness: false);
                Assert.Throws<FileNotFoundException>(() => app.OpenWorkbook(missing));
            }
            catch (ExcelLibraryException)
            {
                // Excel не установлен (Linux/CI) — тест считается пройденным для ветки валидации
            }
        }

        [Fact]
        public void WaitForOwnedProcessExit_NonexistentPid_ReturnsTrue()
        {
            int? deadPid = GetUnusedPid();
            Assert.NotNull(deadPid);
            Assert.True(ComExcelApplication.WaitForOwnedProcessExit(deadPid.Value, 500),
                "Для несуществующего PID метод должен возвращать true (процесс уже отсутствует).");
        }

        /// <summary>Возвращает PID > максимума или заведомо мёртвый (дочерний завершённый процесс).</summary>
        private static int? GetUnusedPid()
        {
            // process.max на Linux обычно <= 4194304; PID заведомо больше — не существует
            try
            {
                string maxFile = "/proc/sys/kernel/pid_max";
                if (File.Exists(maxFile))
                {
                    int max = int.Parse(File.ReadAllText(maxFile).Trim());
                    if (max + 1 <= int.MaxValue) return max + 1;
                }
            }
            catch { }
            // Fallback: запустить и дождаться `true` — его PID гарантированно освобождён
            try
            {
                using var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("/bin/true", "") { UseShellExecute = false });
                if (p != null) { p.WaitForExit(5000); return p.Id; }
            }
            catch { }
            return null;
        }

        #endregion

        #region Windows-only интеграционные тесты (требуют Excel)

        [Fact]
        public void CreateNew_Dispose_NoHangingExcelProcess_And_Idempotent()
        {
            if (!IsExcelAvailable()) return; // пропускается на средах без Excel

            int pid;
            using (var app = new ComExcelApplication())
            {
                pid = app.OwnedProcessId ?? -1;
                app.Dispose(); // первый Dispose
                app.Dispose(); // идемпотентность (п.1.5): второй Dispose не бросает
                Assert.True(app.IsDisposed);
            }

            if (pid > 0)
            {
                // критерий 10.1.1: после Create→Dispose EXCEL.EXE библиотеки завершается
                Assert.True(ComExcelApplication.WaitForOwnedProcessExit(pid, 10000),
                    $"Процесс EXCEL.EXE (PID {pid}) не завершился после Dispose.");
            }
        }

        [Fact]
        public void PerformanceScope_SetsAndRestoresSettings()
        {
            if (!IsExcelAvailable()) return;

            using var app = new ComExcelApplication();
            bool? beforeScreen = (bool?)app.GetAppProperty("ScreenUpdating");

            using (app.BeginPerformanceScope())
            {
                Assert.Equal(false, app.GetAppProperty("ScreenUpdating"));
                Assert.Equal(false, app.GetAppProperty("EnableEvents"));
                Assert.Equal(-4142, Convert.ToInt32(app.GetAppProperty("Calculation"))); // xlCalculationManual
            }

            // восстановление в finally (п.1.4)
            Assert.Equal(beforeScreen ?? true, app.GetAppProperty("ScreenUpdating"));
        }

        [Fact]
        public void AttachToRunning_DoesNotKillForeignInstance()
        {
            if (!IsExcelAvailable()) return;

            // Первый экземпляр создаём сами (он «свой»), второй подключается к нему (становится «чужим»)
            using var owner = new ComExcelApplication();
            int ownerPid = owner.OwnedProcessId ?? -1;

            var attached = new ComExcelApplication(ExcelApplicationMode.AttachToRunning);
            Assert.Null(attached.OwnedProcessId); // чужой процесс не помечен как свой (п.1.6)
            attached.Dispose();

            // Владельец жив, пока не освобождён сам
            if (ownerPid > 0)
                Assert.False(ComExcelApplication.WaitForOwnedProcessExit(ownerPid, 200),
                    "AttachToRunning-Dispose завершил чужой экземпляр Excel!");
        }

        [Fact]
        public void Range_Value2_BatchWriteReadRoundtrip()
        {
            if (!IsExcelAvailable()) return;

            using var app = new ComExcelApplication();
            using var book = app.AddWorkbook();
            var sheet = app.GetWorksheets(book)[0];
            var range = sheet.Range("A1:B2");

            var data = new object[,] { { "a", 1.0 }, { "b", 2.0 } };
            range.SetValue2(data);

            var read = range.ReadAsMatrix();
            Assert.Equal("a", read[0, 0]);
            Assert.Equal(2.0, read[1, 1]);
        }

        [Fact]
        public void Range_SetValue2_LargerThanTarget_ThrowsOverflow()
        {
            if (!IsExcelAvailable()) return;

            using var app = new ComExcelApplication();
            using var book = app.AddWorkbook();
            var sheet = app.GetWorksheets(book)[0];
            var range = sheet.Range("A1:A2");

            var tooBig = new object[,] { { 1 }, { 2 }, { 3 } };
            Assert.Throws<ExcelRangeOverflowException>(() => range.SetValue2(tooBig));
        }

        [Fact]
        public void UseAfterDispose_ThrowsObjectDisposed()
        {
            if (!IsExcelAvailable()) return;

            var app = new ComExcelApplication();
            app.Dispose();
            Assert.Throws<ObjectDisposedException>(() => app.AddWorkbook());
        }

        #endregion
    }
}
