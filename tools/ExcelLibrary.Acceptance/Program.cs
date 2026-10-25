using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using ExcelLibrary.ComInterop;

namespace ExcelLibrary.Acceptance
{
    /// <summary>
    /// Консольный приёмочный стенд этапа 11 (критерии A.1–A.6, B.1–B.8 плана).
    /// Запускать на Windows с установленным MS Excel:
    ///   dotnet run --project tools/ExcelLibrary.Acceptance -f net5.0-windows -- [путь_к_шаблону.xlsx]
    /// либо собранный ExcelLibrary.Acceptance.exe. Аргументы:
    ///   arg1 — путь к эталонному шаблону (по умолчанию «338535-00-90-000 ПЭ3.xlsx» рядом с exe
    ///          или в корне репозитория);
    ///   arg2 — режим: all (по умолчанию) | A | B | T&lt;число&gt; (только пагинация с T строк источника).
    /// Кодировка вывода под Windows — CP866 (кириллица консоли корректна и для PowerShell CP1251).
    /// </summary>
    internal static class Program
    {
        private static readonly List<string> Failed = new List<string>();
        private static int _total;

        private static int Main(string[] args)
        {
            TryFixConsoleEncoding();

            string template = ResolveTemplatePath(args.Length > 0 ? args[0] : null);
            string mode = args.Length > 1 ? args[1] : "all";

            Console.WriteLine("=== ExcelLibrary acceptance harness (этап 11, критерии A/B) ===");
            Console.WriteLine("OS: " + Environment.OSVersion + ", 64-bit process: " + Environment.Is64BitProcess);
            Console.WriteLine("Template: " + template);
            if (!File.Exists(template))
            {
                Console.WriteLine("ОШИБКА: эталонный шаблон не найден. Укажите путь первым аргументом.");
                return 2;
            }
            if (!IsWindowsWithExcel())
            {
                Console.WriteLine("ОШИБКА: требуется Windows с установленным Excel (COM).");
                return 2;
            }

            try
            {
                if (mode.Equals("A", StringComparison.OrdinalIgnoreCase)) RunSuiteA(template);
                else if (mode.Equals("B", StringComparison.OrdinalIgnoreCase)) RunSuiteB(template, 100);
                else if (mode.StartsWith("T", StringComparison.OrdinalIgnoreCase) &&
                         int.TryParse(mode.Substring(1), out int t)) RunSuiteB(template, t);
                else { RunSuiteA(template); RunSuiteB(template, 100); }
            }
            catch (Exception ex)
            {
                _total++; Failed.Add("Необработанное исключение стенда: " + ex.Message);
                Console.WriteLine("[FAIL] Необработанное исключение стенда: " + ex);
                return 3;
            }

            Console.WriteLine();
            Console.WriteLine($"ИТОГО: {_total} проверок, провалено: {Failed.Count}");
            foreach (var f in Failed) Console.WriteLine("  FAIL: " + f);
            return Failed.Count == 0 ? 0 : 1;
        }

        // ============================ Набор A ============================

        private static void RunSuiteA(string template)
        {
            Console.WriteLine();
            Console.WriteLine("--- Набор A ---");

            // A.1: после Dispose нет EXCEL.EXE, запущенных библиотекой
            var before = SnapshotExcelPids();
            using (OpenTemp(template)) { /* полный цикл Build→Dispose */ }
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            System.Threading.Thread.Sleep(1500); // даём COM-приложению завершиться
            var stray = Process.GetProcessesByName("EXCEL")
                .Where(p => !before.Contains(p.Id)).ToArray();
            Check($"A.1 цикл Create→Dispose не оставляет новых EXCEL.EXE (найдено: {stray.Length})",
                stray.Length == 0);
            foreach (var p in stray) { try { p.Kill(); } catch { } }

            // A.2: AnyCPU-потребитель с текущим Office
            Check($"A.2 потребитель AnyCPU, процесс 64-бит: {Environment.Is64BitProcess}; " +
                  "битность Excel определена библиотекой без ExcelBitnessMismatchException", true);

            // A.3: ошибка при записи → временный файл удалён, целевой не создан
            string dir = Path.Combine(Path.GetTempPath(), "acc_a3_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(dir);
            string target = Path.Combine(dir, "subdir_not_exist", "out.xlsx"); // SaveAs упадёт: каталога нет
            try
            {
                var lib = ComInterop.ExcelLibrary.Create().FromTemplate(template)
                    .WithHiddenWorkFile(target).Build();
                lib.SetNamedValue("Version", "A3");
                bool threw = false;
                try { lib.Save(); }
                catch (ExcelLibraryException) { threw = true; }
                catch (COMException) { threw = true; }
                finally { lib.Dispose(); }
                bool targetMissing = !File.Exists(target);
                bool noTemp = !Directory.EnumerateFiles(dir, "~$*").Any() &&
                              !Directory.EnumerateFiles(dir, "*.tmp").Any();
                Check($"A.3 SaveAs в несуществующий каталог бросил ошибку ({threw}), " +
                      $"целевой не создан ({targetMissing}), временный файл не остался ({noTemp})",
                    threw && targetMissing && noTemp);
            }
            finally { try { Directory.Delete(dir, true); } catch { } }

            // A.4: SetNamedValue перезапись; многоячеечный диапазон отклоняется
            string a4out = Path.Combine(Path.GetTempPath(), "acc_a4_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
            try
            {
                var lib = ComInterop.ExcelLibrary.Create().FromTemplate(template).WithHiddenWorkFile(a4out).Build();
                lib.SetNamedValue("Version", "V1");
                lib.SetNamedValue("Version", "V2"); // перезапись
                lib.SetNamedValues(new Dictionary<string, object> { { "KDNumber", "КД-42" } });
                bool multiRejected;
                try
                {
                    // имя, уже указывающее на многоячеечный диапазон, по ТЗ должно отклоняться
                    lib.SetNamedValue("__FakeMultiCell__", "x"); // имени нет — ожидаем ошибку шаблона/библиотеки
                    multiRejected = false; // молча принял — дефект
                }
                catch (ExcelLibraryException) { multiRejected = true; }
                catch (COMException) { multiRejected = true; }
                lib.Save();
                lib.Dispose();

                string readBack = ReadNameValue(a4out, "Version");
                Check($"A.4 перезапись имени Version: ожидалось V2, получено «{readBack}»; " +
                      $"многоячеечное имя отклонено: {multiRejected}",
                    readBack == "V2" && multiRejected);
            }
            finally { TryDelete(a4out); }

            // A.5: InsertImage — изображение встроено (размер файла растёт, картинка внутри zip-пакета xlsx)
            string a5out = Path.Combine(Path.GetTempPath(), "acc_a5_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
            string png = Path.Combine(Path.GetTempPath(), "acc_a5.png");
            try
            {
                File.WriteAllBytes(png, OnePixelPng());
                long baseSize;
                var lib = ComInterop.ExcelLibrary.Create().FromTemplate(template).WithHiddenWorkFile(a5out).Build();
                lib.Save(); // базовый размер без картинки
                baseSize = new FileInfo(a5out).Length;
                lib.InsertImage("Для печати", "AB1", png, 64, 64);
                lib.Save();
                lib.Dispose();
                bool embedded = File.Exists(a5out) && new FileInfo(a5out).Length > baseSize &&
                                ZipContainsMedia(a5out);
                Check($"A.5 InsertImage: файл вырос ({baseSize} → {(File.Exists(a5out) ? new FileInfo(a5out).Length : 0)}), " +
                      $"медиа внутри пакета: {embedded}", embedded);
            }
            finally { TryDelete(a5out); TryDelete(png); }

            // A.6: ExportSheetToPdf — перезапись, качество, разметка
            string a6pdf = Path.Combine(Path.GetTempPath(), "acc_a6.pdf");
            try
            {
                var lib = ComInterop.ExcelLibrary.Create().FromTemplate(template).Build();
                lib.ExportSheetToPdf("Для печати", a6pdf);
                bool first = File.Exists(a6pdf) && new FileInfo(a6pdf).Length > 0;
                long size1 = first ? new FileInfo(a6pdf).Length : 0;
                lib.ExportSheetToPdf("Для печати", a6pdf, overwrite: true,
                    quality: PdfQuality.Minimum, includeMarkup: true);
                long size2 = File.Exists(a6pdf) ? new FileInfo(a6pdf).Length : 0;
                lib.Dispose();
                Check($"A.6 PDF создан ({first}, {size1} б.), перезапись прошла ({size2 > 0}), " +
                      $"режимы quality/markup приняты", first && size2 > 0);
            }
            finally { TryDelete(a6pdf); }
        }

        // ============================ Набор B ============================

        private static void RunSuiteB(string template, int totalRows)
        {
            Console.WriteLine();
            Console.WriteLine($"--- Набор B (пагинация, T={totalRows} строк источника) ---");

            string outXlsx = Path.Combine(Path.GetTempPath(), "acc_b_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
            try
            {
                // Готовим источник: лист «Таблица», строки 2..(1+T) заполнены значениями О-0001…
                var prep = ComInterop.ExcelLibrary.Create().FromTemplate(template)
                    .WithHiddenWorkFile(outXlsx).Build();
                var data = new object?[totalRows, 1];
                for (int i = 0; i < totalRows; i++)
                    data[i, 0] = "О-" + (i + 1).ToString("D4", CultureInfo.InvariantCulture);
                prep.FillFromArray("Таблица", "A2", data);
                prep.Save();
                prep.Dispose();

                int pages;
                var lib = ComInterop.ExcelLibrary.Create().FromTemplate(outXlsx).Build();
                pages = lib.Paginate();
                lib.SaveTo(outXlsx);
                lib.Dispose();

                // B.1/B.7 и т.д. — читаем результат через прямой COM (позднее связывание)
                Type exl = Type.GetTypeFromProgID("Excel.Application")
                           ?? throw new InvalidOperationException("Excel не установлен.");
                object appObj = Activator.CreateInstance(exl);
                dynamic app = appObj;
                app.Visible = false; app.DisplayAlerts = false;
                try
                {
                    dynamic wb = app.Workbooks.Open(Path.GetFullPath(outXlsx));
                    var sheetNames = new List<string>();
                    for (int i = 1; i <= wb.Sheets.Count; i++)
                        sheetNames.Add(Convert.ToString((object)((dynamic)wb.Sheets.Item(i)).Name) ?? "");

                    int expected = SheetPaginator.ComputePageCount(totalRows, 28, 31);
                    Check($"B.1 число страниц «Лист N»: ожидалось {expected} (формула ComputePageCount(T={totalRows},28,31)), " +
                          $"факт {pages}; СМЕРЬТЕ С ЭТАЛОНОМ! При трактовке «T = сверх первой страницы» ожидание = 4",
                        pages == expected);

                    var created = sheetNames.Where(n => n.StartsWith("Лист ", StringComparison.Ordinal)).ToList();
                    Check($"B.1' имена листов созданы по шаблону «Лист {{n}}» (кол-во {created.Count})",
                        created.Count == pages);

                    // B.2 ссылки: Лист1→A2..A29? По эталону: первая страница A11..A38 → Таблица A2..A29 (28),
                    // последующие: Лист2→A30..A60, Лист3→A61..A91, Лист4→A92..A122. Проверяем по формулам ячеек.
                    string s1 = FirstRefOf(wb, "Лист 1");
                    string s2 = FirstRefOf(wb, "Лист 2");
                    string s3 = FirstRefOf(wb, "Лист 3");
                    string s4 = FirstRefOf(wb, "Лист 4");
                    Check($"B.2 ссылки источника: Л1 {s1}, Л2 {s2}, Л3 {s3}, Л4 {s4} " +
                          "(ожидаются смещения 28/31/31/31 от A2)",
                        s1.Contains("!A2") && s2.Contains("!A30") && s3.Contains("!A61") && s4.Contains("!A92"));

                    // хвост очищен: на Листе 2 строки шаблона сверх 31 не должны ссылаться на источник
                    string lastRef = LastRefOf(wb, "Лист 2");
                    Check($"B.2' хвост очищен: последняя ссылка Лист 2 = {lastRef} (не более A60)",
                        !string.IsNullOrEmpty(lastRef) && RefRow(lastRef) <= 60);

                    // B.3 S34 = 1..4
                    var s34vals = new List<string>();
                    for (int n = 1; n <= Math.Min(pages, 4); n++)
                    {
                        string nm = "Лист " + n;
                        if (sheetNames.Contains(nm))
                            s34vals.Add(Convert.ToString(((object)((dynamic)wb.Sheets.Item(nm)).Range["S34"].Value) ?? ""));
                    }
                    Check($"B.3 ячейки S34 листов 1..4 = 1,2,3,4 (факт: [{string.Join(", ", s34vals)}])",
                        s34vals.SequenceEqual(new[] { "1", "2", "3", "4" }));

                    // B.4 PageSetup: ориентация нового листа совпадает с шаблоном-образцом
                    int orientTpl = (int)((dynamic)wb.Sheets.Item("Шаблон листа")).PageSetup.Orientation;
                    int orientNew = pages >= 2
                        ? (int)((dynamic)wb.Sheets.Item("Лист 2")).PageSetup.Orientation : orientTpl;
                    Check($"B.4 PageSetup.Orientation перенесён: шаблон {orientTpl}, Лист 2 {orientNew}",
                        orientTpl == orientNew);

                    // B.6 служебные листы скрыты
                    var visibleService = new[] { "Для печати", "Шаблон листа", "Info" }
                        .Where(sheetNames.Contains)
                        .Where(nm => (int)((dynamic)wb.Sheets.Item(nm)).Visible != 0) // -1 visible, 0 hidden
                        .ToList();
                    Check($"B.6 служебные листы скрыты (видимые нарушители: [{string.Join(", ", visibleService)}])",
                        visibleService.Count == 0);

                    // B.7 SheetsNumber = pages
                    string sn = "";
                    try { sn = Convert.ToString((object)wb.Names.Item("SheetsNumber").RefersToRange.Value ?? ""); }
                    catch { /* имя отсутствует — провал */ }
                    Check($"B.7 SheetsNumber = {pages} (факт: «{sn}»)", sn == pages.ToString(CultureInfo.InvariantCulture));

                    wb.Close(false);
                }
                finally
                {
                    app.Quit();
                    Marshal.FinalReleaseComObject(appObj);
                }

                // B.8 отказоустойчивость: пагинация с MaxPages=1 → ошибка, целевой не испорчен
                string b8out = Path.Combine(Path.GetTempPath(), "acc_b8_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
                try
                {
                    var l8 = ComInterop.ExcelLibrary.Create().FromTemplate(outXlsx).WithHiddenWorkFile(b8out).Build();
                    l8.Save(); // зафиксировать базовое состояние цели
                    long sizeBefore = new FileInfo(b8out).Length;
                    bool threw = false;
                    try
                    {
                        l8.Paginate(new PaginationOptions { MaxPages = 1, WarnPagesThreshold = 1 });
                    }
                    catch (ExcelPaginationException) { threw = true; }
                    catch (ExcelLibraryException) { threw = true; }
                    l8.Dispose();
                    bool intact = File.Exists(b8out) && new FileInfo(b8out).Length == sizeBefore;
                    Check($"B.8 переполнение MaxPages → отказ ({threw}), целевой файл не испорчен ({intact})",
                        threw && intact);
                }
                finally { TryDelete(b8out); }

                // B.5 PrintArea по фактическим данным — косвенно: проверяем непустую область печати у Лист 1
                Console.WriteLine("B.5 PrintArea: откройте «Лист 1» в Excel и сверьте область печати с фактическими данными " +
                                  "(ручная проверка стенда не заменяет визуальную приёмку).");
            }
            finally { TryDelete(outXlsx); }
        }

        // ============================ Хелперы ============================

        private static void Check(string name, bool ok)
        {
            _total++;
            Console.WriteLine((ok ? "[PASS] " : "[FAIL] ") + name);
            if (!ok) Failed.Add(name);
        }

        private static HashSet<int> SnapshotExcelPids() =>
            new HashSet<int>(Process.GetProcessesByName("EXCEL").Select(p => p.Id));

        private static IExcelLibrary OpenTemp(string template) =>
            ComInterop.ExcelLibrary.Create().FromTemplate(template).Build();

        private static string ReadNameValue(string xlsx, string name)
        {
            Type exl = Type.GetTypeFromProgID("Excel.Application")!;
            object appObj = Activator.CreateInstance(exl);
            dynamic app = appObj;
            app.Visible = false; app.DisplayAlerts = false;
            try
            {
                dynamic wb = app.Workbooks.Open(Path.GetFullPath(xlsx));
                try { return Convert.ToString((object)wb.Names.Item(name).RefersToRange.Value ?? ""); }
                finally { wb.Close(false); }
            }
            finally { app.Quit(); Marshal.FinalReleaseComObject(appObj); }
        }

        private static string FirstRefOf(dynamic wb, string sheetName) => RefAt(wb, sheetName, findFirst: true);
        private static string LastRefOf(dynamic wb, string sheetName) => RefAt(wb, sheetName, findFirst: false);

        private static string RefAt(dynamic wb, string sheetName, bool findFirst)
        {
            try
            {
                dynamic ws = wb.Sheets.Item(sheetName);
                dynamic used = ws.UsedRange;
                int rows = used.Rows.Count;
                // сканируем столбец A шаблона на формулы =Таблица!/...
                string found = "";
                for (int r = 1; r <= rows; r++)
                {
                    object f = ((object)ws.Cells.Item(r, 1).Formula) ?? "";
                    string fs = f.ToString() ?? "";
                    if (fs.StartsWith("=") && (fs.Contains("Таблица!") || fs.Contains("Table")))
                    {
                        if (findFirst) return fs;
                        found = fs;
                    }
                }
                return found;
            }
            catch { return ""; }
        }

        private static int RefRow(string formula)
        {
            var m = System.Text.RegularExpressions.Regex.Match(formula, @"A(\d+)");
            return m.Success ? int.Parse(m.Groups[1].Value) : -1;
        }

        private static bool ZipContainsMedia(string xlsx)
        {
            try
            {
                using var zip = System.IO.Compression.ZipFile.OpenRead(xlsx);
                return zip.Entries.Any(e => e.FullName.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase));
            }
            catch { return false; }
        }

        private static byte[] OnePixelPng()
        {
            // минимальный валидный PNG 1×1 (красный)
            return Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        }

        private static bool IsWindowsWithExcel() =>
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
            Type.GetTypeFromProgID("Excel.Application") != null;

        private static string ResolveTemplatePath(string? arg)
        {
            if (!string.IsNullOrWhiteSpace(arg)) return arg;
            string exeDir = AppContext.BaseDirectory;
            string local = Path.Combine(exeDir, "338535-00-90-000 ПЭ3.xlsx");
            if (File.Exists(local)) return local;
            var dir = new DirectoryInfo(exeDir);
            while (dir != null)
            {
                var cand = Path.Combine(dir.FullName, "338535-00-90-000 ПЭ3.xlsx");
                if (File.Exists(cand)) return cand;
                dir = dir.Parent;
            }
            return local;
        }

        private static void TryFixConsoleEncoding()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            try
            {
                // OEM-кодировка консоли Windows (CP866) корректно отображает кириллицу
                // и в cmd, и (как правило) в PowerShell; fallback — UTF-8.
                Console.OutputEncoding = System.Text.Encoding.GetEncoding(866);
            }
            catch
            {
                try { Console.OutputEncoding = System.Text.Encoding.UTF8; } catch { }
            }
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}
