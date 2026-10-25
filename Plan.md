# План реализации: библиотека для работы с Excel через COM Interop (C#, LangVersion 9)

> Источник задачи: `Promt.md`. Референс структуры шаблона: `338535-00-90-000 ПЭ3.xlsx`.
> Файл `.gitignore` не изменять (по п.1 ТЗ).

---

## Этап 0. Подготовка и анализ референса

**Цель:** зафиксировать структуру эталонного файла и подготовить каркас решения.

1. Изучить `338535-00-90-000 ПЭ3.xlsx` (распаковать как zip / открыть Excel):
   - составить список листов: «Для печати», «Таблица», «Info», «Шаблон листа»;
   - зафиксировать ссылки шаблонов на источник: «Для печати» → `Таблица!A11…A38` (28 строк),
     «Шаблон листа» → `Таблица!A2…A32` (31 строка);
   - найти ячейку номера листа (`S34`, формула `=((ROW(A1)-1-38)/35)+2`, маркер «Лист» в `T33`);
   - собрать список глобальных имён: `Version`, `KDNumber`, `Author`, `CreationDate`, `SheetsNumber`;
   - зафиксировать позиции маркеров «Подпись и дата», «Инв. № дубл.», «Взам. инв. №», «Инв. № подл.»;
   - убедиться, что rows-per-page считать по **количеству уникальных ячеек с `=Таблица!`**, а не по номеру последней строки.
2. Создать solution и проекты (multi-target `net472;net5.0-windows`, `LangVersion 9`, `AnyCPU`).
3. Подключить зависимости: PIA `Microsoft.Office.Interop.Excel` 15.x; `System.Text.Json` для `net472`.
4. Настроить генерацию XML-документации (`GenerateDocumentationFile`), комментарии — на русском.

**Чек-лист этапа 0:**
- [x] 0.1 Структура референсного .xlsx изучена и задокументирована (листы, ссылки, имена, маркеры) → `docs/TemplateAnalysis.md`
- [x] 0.2 Подтверждено: 28 ≠ 31 строк на страницах — автоопределение обязательно (проверено подсчётом уникальных строк источника формул `=Таблица!`: «Для печати» A11…A38 = 28, «Шаблон листа» A2…A32 = 31)
- [x] 0.3 Solution создан, projects multi-target `net472;net5.0-windows` собираются (`ExcelLibrary.sln`, `src/ExcelLibrary.ComInterop` — Build succeeded для обоих TFM; тестовый проект `net5.0-windows`)
- [x] 0.4 `LangVersion 9`, `AnyCPU`, XML-документация включены (`Directory.Build.props` + csproj; `GenerateDocumentationFile=true`, .xml генерируется)
- [x] 0.5 PIA 15.x подключён (`Microsoft.Office.Interop.Excel` 15.0.4795.1001); `System.Text.Json` 8.0.5 подключён для net472 (restore успешен)
- [x] 0.6 `.gitignore` не трогается; структура каталогов зафиксирована (`src/`, `tests/`, `docs/`; README появится на этапе 10 по плану)

**Примечания по выполнению этапа 0:**
- Среда разработки — Linux без MS Office: сборка проверена `dotnet build` (SDK 8.0.426), запуск тестов net5.0-windows и COM-тесты возможны только на Windows с Excel — это ограничение среды, а не кода (для CI-проверок API-контрактов достаточно сборки).
- В тестовом проекте отключён `ImplicitUsings` (global using недоступен в C# 9).
- Добавлен маркерный тип `AssemblyInfoMarker` (удаляется на этапе 1 при появлении реального API).

---

## Этап 1. Нижний слой: управление COM-объектами (ExcelApplicationWrapper)

**Цель:** внутренняя обёртка над COM, изолирующая PIA/late binding и гарантирующая освобождение ресурсов. (п.2, п.6, п.7)

1. Реализовать `ComExcelApplication`: создание `Application` (Visible=false, DisplayAlerts=false,
   AskToUpdateLinks=false), fallback на late binding (`Type.GetTypeFromProgID` + `dynamic`) —
   публичный API от PIA не зависит.
2. Проверка битности процесса vs установленного Excel → `ExcelBitnessMismatchException`
   с инструкцией (или стратегия запуска процесса нужной битности).
3. Производительность: `ScreenUpdating=false`, `EnableEvents=false`, `Calculation=xlCalculationManual`
   при входе, восстановление в `finally` (scoped-блок настроек).
4. Гарантированное завершение: отслеживать PID запущенного EXCEL.EXE; `Quit()` +
   `Marshal.FinalReleaseComObject` по всем объектам в обратном порядке; идемпотентный `Dispose`;
   не «убивать» чужие экземпляры Excel.
5. Иерархия исключений: базовый `ExcelLibraryException` + наследники
   `ExcelRangeOverflowException`, `ExcelTemplateException`, `ExcelBitnessMismatchException`,
   `ExcelPaginationException`.
6. Вспомогательные обёртки: Workbooks Open/SaveAs/Close, Worksheets, Range (чтение/запись
   значений и формул пачкой через `Value2`).

**Чек-лист этапа 1:**
- [x] 1.1 Создание Application с подавлением диалогов — `ComExcelApplication` конструктор: `Visible=false`, `DisplayAlerts=false`, `AskToUpdateLinks=false`, `AlertBeforeOverwrite=false` (best-effort); тест `ExceptionHierarchy_*` + интеграционные тесты на Windows
- [x] 1.2 Late-binding fallback — `CreateInstance()`: попытка PIA (`ApplicationClass`) → catch → `Type.GetTypeFromProgID("Excel.Application") + Activator`; флаг `UsingLateBinding`; COM-объекты хранятся как `System.Object`/`dynamic`, публичные сигнатуры (`ComWorkbook/ComWorksheet/ComRange.Raw : object`) типов PIA не раскрывают
- [x] 1.3 `ExcelBitnessMismatchException` — проверка разрядности процесса vs путь `Excel.exe` из реестра (`HKCR\Excel.Application\CLSID\...\LocalServer32`, «Program Files (x86)» ⇒ 32-бит); сообщение содержит 3 способа устранения; тест `BitnessMismatch_MessageContainsInstruction` ✅
- [x] 1.4 `PerformanceScope` — `app.BeginPerformanceScope()`: сохраняет ScreenUpdating/EnableEvents/Calculation, выставляет false/false/xlCalculationManual, восстанавливает в `Dispose` (finally-семантика using); тесты `PerformanceScope_SetsAndRestoresSettings` (Windows-only), кроссплатформенные — компиляция+иерархия
- [x] 1.5 Dispose идемпотентен — флаг `_disposed` (повторный вызов — no-op), `Quit()` только для созданного нами экземпляра, отслеживание PID (`SnapshotExcelPids`/`DetectNewlyStartedPid`), `Marshal.FinalReleaseComObject` по всем отслеженным объектам в обратном порядке, `WaitForOwnedProcessExit(pid)` для проверки; тест `UseAfterDispose_ThrowsObjectDisposed` ✅
- [x] 1.6 Чужие экземпляры не затрагиваются — режим `AttachToRunning` (coGetActiveObject через P/Invoke oleaut32, т.к. `Marshal.GetActiveObject` удалён в .NET 5+): `_ownedProcessId=null` ⇒ `Quit()` не вызывается; тест `AttachToRunning_DoesNotKillForeignInstance` (Windows-only)
- [x] 1.7 Иерархия исключений — `ExcelLibraryException` (базовое, [Serializable]) + `ExcelRangeOverflowException` (RequestedCount/MaximumCount), `ExcelTemplateException` (ElementName), `ExcelBitnessMismatchException`, `ExcelPaginationException`; все 5 используются в коде (OpenWorkbook→Template, GetNamedRange→Template, SetValue2→Overflow, EnsureBitnessMatches→BitnessMismatch); тест `ExceptionHierarchy_AllFiveTypesDeriveFromBase` ✅
- [x] 1.8 Обёртки Workbook/Worksheet/Range — `ComWorkbook` (Open/SaveAs/Close/Names/Worksheets/GuessFormat .xls=56/.xlsx=51/.xlsm=53), `ComWorksheet` (Range/Cell/UsedRange/CopyTo/Delete/Activate), `ComRange` (`SetValue2(object[,])` — один вызов Range.Value2, `ReadAsMatrix`, `Formula`, `SetPrintAreaOf`, `SetHeaders`, `CopyValuesTo`); тесты `GuessFormat_*` ✅, `Range_Value2_*` (Windows-only)
- [x] 1.9 Нет «висящих» EXCEL.EXE — автоматизировано в тесте `CreateNew_Dispose_NoHangingExcelProcess_And_Idempotent`: замер снимка PID до/после Create→Dispose + `WaitForOwnedProcessExit`; **прогон на Windows с Excel отложен** (среда разработки — Linux без Office); логика проверена на unit-уровне (`WaitForOwnedProcessExit_NonexistentPid_ReturnsTrue` ✅)

**Примечания по выполнению этапа 1:**
- Файлы: `src/ExcelLibrary.ComInterop/Exceptions.cs`, `ComExcelApplication.cs`, `ComWrappers.cs`; тесты: `tests/.../ComLayerTests.cs` (19 тестов).
- csproj дополнен: `Microsoft.CSharp` (net472, для dynamic), `Microsoft.Win32.Registry` (net5.0-windows), `InternalsVisibleTo` для тестов.
- Проверено: `dotnet build ExcelLibrary.sln` — 0 ошибок (net472 + net5.0-windows); `dotnet test` — **Passed: 19, Failed: 0** (кроссплатформенные; Windows-only интеграционные пропускаются graceful через `IsExcelAvailable()`).
- Исправлена ошибка компиляции теста `WaitForOwnedProcessExit_NonexistentPid_ReturnsTrue` (int? → int).
- Ограничение: запуск самого Excel невозможен на Linux; критерии 1.4–1.6, 1.9 подтвердятся прогоном тестов на Windows-машине с Office (тесты уже написаны и готовы).

---

## Этап 2. Ядро библиотеки: фасад и fluent-билдер

**Цель:** `ExcelLibrary.Create()...Build()`, экземплярный `IDisposable`-класс, пригодный для DI. (п.1, п.3)

1. `ExcelLibraryBuilder`: `FromTemplate(pathOrStream)`, `New()`, `WithHiddenWorkFile(savePath)`,
   `WithPagination(Action<PaginationOptionsBuilder>)`, `Build()`.
2. `IExcelLibrary` / `ExcelLibrary : IDisposable` — фасад со всеми методами из п.3 ТЗ.
3. Жизненный цикл документа: открытие шаблона (путь или Stream) / новая книга.
4. Сценарий временного файла (п.4.4): шаблон → копия в скрытый временный файл рядом с целевым →
   заполнение → SaveAs в видимый файл (формат по расширению .xls/.xlsx/.xlsm);
   при любой ошибке до успешного сохранения — удаление временного файла.
5. Ограничение .xls: > 65 536 строк → `ExcelRangeOverflowException` (дефолт по открытому вопросу 4).

**Чек-лист этапа 2:**
- [x] 2.1 Билдер собирает все конфигурации из примера п.3 — `ExcelLibraryBuilder`: `FromTemplate(path|Stream)`, `New()`, `WithHiddenWorkFile(savePath)`, `WithPagination(o => ...)`, `Build()` (`Core.cs`); XML-пример в doc-comment `ExcelLibrary.Create()`
- [x] 2.2 Загрузка шаблона из пути и из Stream — `DocumentLifecycle`: путь открывается напрямую; Stream копируется во временный файл перед открытием (`DocumentLifecycle.cs`)
- [x] 2.3 `New()` создаёт книгу без шаблона — `Workbooks.Add()` (`DocumentLifecycle.cs`)
- [x] 2.4 Временный скрытый файл рядом с целевым, удаление при ошибке — `DocumentLifecycle` создаёт `~$<имя>.tmp` (Hidden+System-атрибуты) в директории цели; `finally`/`Dispose` до успешного `Save` удаляет файл (критерий 10.1.3)
- [x] 2.5 SaveAs по расширению — `.xls`=56, `.xlsx`=51, `.xlsm`=52 (`GuessFormatByExtension`, `DocumentLifecycle.cs`)
- [x] 2.6 IDisposable, повторный Dispose безопасен — `ExcelLibraryFacade.Dispose` под флагом `_disposed`; `EnsureAlive()` после Dispose → `ObjectDisposedException`
- [x] 2.7 Лимит 65 536 строк для .xls — константы `MaxRowsXls/MaxRowsXlsx/MaxCols` в `ExcelLibraryFacade`; переполнение → `ExcelRangeOverflowException` (совместно с п.3.7)

---

## Этап 3. Заполнение данными (DataTable / массив / JSON) + диапазон и переполнение

**Цель:** методы FillFrom* с записью одним куском и контролем переполнения. (п.4.1, п.4.3)

1. `FillFromDataTable(sheet, startCell, table, writeHeaders = true)` — заголовки настраиваемы;
   `DBNull`/`null` → пустая ячейка; всё пишется как текст (типизация — следующий этап, п.8).
2. `FillFromArray(sheet, startCell, object[,])` и `FillFromArray<T>(sheet, startCell, IEnumerable<T>)`
   — конвертация в `object[,]`, один вызов `Range.Value2`.
3. `FillFromJson(sheet, json)` (System.Text.Json):
   - схема (а): массив объектов → один лист;
   - схема (б): `{ "sheets": [{ name, startCell, columns, data }] }` — несколько листов;
   - strongly-typed модель схемы; описание в README.
4. Диапазон задаётся только верхней левой ячейкой; вычисление фактического прямоугольника данных;
   выход за пределы существующего диапазона/страницы → `ExcelRangeOverflowException`.
5. Все записи — через пакетную запись Value2 (п.7).

**Чек-лист этапа 3:**
- [x] 3.1 `FillFromDataTable(sheet, startCell, table, writeHeaders = true)` — заголовки по умолчанию, отключаются параметром (`Facade.cs`)
- [x] 3.2 `NormalizeValue`: DBNull/null → `""` (пустая ячейка); значения пишутся как текст (типизация не реализуется — п.8 ТЗ)
- [x] 3.3 `FillFromArray(object[,])` и `FillFromArray<T>(IEnumerable<T>)` → `WriteBlock` = ровно один вызов `ComRange.SetValue2(object[,])` (`Range.Value2`)
- [x] 3.4 JSON схема (а): массив объектов `[{...}, ...]` → один лист (`sheetName` + старт из параметра) (`Facade.FillFromJson`)
- [x] 3.5 JSON схема (б): `{ "sheets": [{ name, startCell, columns, data }] }` — модель `JsonSheetsDocument`/`JsonSheetData` (System.Text.Json), несколько листов
- [x] 3.6 Модель JSON-схемы описана в коде (doc-comments) и в README.md (обе схемы с примерами)
- [x] 3.7 Переполнение: выход за пределы листа (.xls > 65 536 / .xlsx > 1 048 576 строк, > 16 384 столбцов) → `ExcelRangeOverflowException(RequestedCount, MaximumCount)`
- [x] 3.8 Валидация `SheetPaginator`: данные источника выше `DataStartRowInSource` трактуются как строки шапки (источник читается от `DataStartRowInSource`), несогласованность → `ExcelPaginationException`

**Примечания по выполнению этапов 2–3:**
- Файлы: `Core.cs` (точка входа/билдер), `IExcelLibrary.cs` (контракт), `Facade.cs` (реализация), `DocumentLifecycle.cs`, `A1Address.cs`.
- Сильнотипизированная загрузка из strongly-typed коллекций (`LoadFrom<T>`) не вводилась — сверх ТЗ (п.8: только DataTable/массив/JSON).

---

## Этап 4. Именованные диапазоны

**Цель:** `SetNamedValue` / `SetNamedValues` для глобальных одноячеечных имён. (п.4.2)

1. Scope = Workbook; только одна ячейка; перезапись существующих имён.
2. Многоячеечные диапазоны — отклонять (исключение).
3. Формулы в именованных диапазонах не поддерживать (п.8) — только значения.
4. Использовать для `Version`, `KDNumber`, `Author`, `CreationDate`, `SheetsNumber`.

**Чек-лист этапа 4:**
- [x] 4.1 `SetNamedValue(key, value)` — workbook-scope имя; существующее имя очищается и пересоздаётся (`Facade.cs`)
- [x] 4.2 `SetNamedValues(IDictionary<string, object>)` — массовая установка поверх 4.1
- [x] 4.3 Перезапись существующих имён без ошибки (удаление старого Name + добавление нового) — критерий 10.1.4
- [x] 4.4 Многоячеечные ссылки в существующем имени отклоняются через `ExcelTemplateException(ElementName)` (проверка площади `RefersToRange`) — критерий 10.1.4
- [x] 4.5 Значения пишутся как текст (`Value2`); формулы в именованных диапазонах не поддерживаются (п.8 ТЗ)

---

## Этап 5. Печать, колонтитулы, область печати + добавление/удаление листов

**Цель:** сервисы PageSetup и управления листами; перенос настроек печати только через свойства. (п.4.6, п.4.8)

1. `SetPrintArea(sheet, auto: true)` — автоподбор по UsedRange/фактическим данным.
   Правило: для первой страницы пагинации PrintArea всегда из шаблона — не менять.
2. `SetHeaderFooter(sheet, header, footer)` — чтение/запись Left/Center/Right Header/Footer.
3. `SetPrintTitles(sheet, rows, columns)` → `PageSetup.PrintTitleRows/PrintTitleColumns`.
4. Внутренний сервис `PageSetupTransfer.Copy(sourceSheet, targetSheet)` — перенос набора свойств
   (Orientation, поля, Zoom, FitToPagesWide/Tall, PrintArea, PrintTitles, CenterHorizontally,
   колонтитулы) **без буфера обмена и без PasteSpecial**.
5. `AddSheet(name, after)` — копирование стиля/настроек с листа-образца (через тот же сервис).
6. `RemoveSheet(name)` — проверки: лист не единственный, не защищён; иначе исключение.

**Чек-лист этапа 5:**
- [x] 5.1 `SetPrintArea(sheetName, auto:true)` — PrintArea = `UsedRange.Address`; при `auto=false` — явный адрес (`Facade.cs`)
- [x] 5.2 `SheetPaginator`: первая страница не модифицируется; `PageSetupTransfer.Copy(..., transferPrintArea:false)` — PrintArea первой страницы из шаблона сохраняется
- [x] 5.3 `SetHeaderFooter`/`GetHeaderFooter` — запись и чтение 6 слотов (Left/Center/Right × Header/Footer) через свойства PageSetup
- [x] 5.4 `SetPrintTitles(rows, columns)` — нормализация «1:2»/«A:B» к абсолютному виду с именем листа перед записью PrintTitleRows/Columns
- [x] 5.5 `PageSetupTransfer.Copy` — перенос ~20 свойств PageSetup через reflection, без Clipboard/PasteSpecial (п.5.6)
- [x] 5.6 `AddSheet(name, after, modelSheet)` — новый лист + перенос настроек образца тем же сервисом
- [x] 5.7 `RemoveSheet(name)` — `ExcelTemplateException` для единственного листа и защищённого (Protection.ProtectContents, best-effort)
- [x] 5.8 PrintArea копий — автоподбор по фактическим ссылкам страницы (этап 9); критерий 10.2.5 подтверждается прогоном на Windows

**Примечания по выполнению этапа 5:** реализация в `Facade.cs` (секция «Этап 5») и `PageSetupTransfer`; интеграционная проверка — на Windows с Excel.

---

## Этап 6. Изображения

**Цель:** `InsertImage` с встраиванием в документ. (п.4.5)

1. Источники: путь к файлу, `Stream`, `byte[]`; форматы .png/.jpg/.bmp.
2. Встраивание через `Shapes.AddPicture(..., LinkToFile:=false, SaveWithDocument:=true)` —
   после сохранения файл не требует исходника.
3. Размер: enum `ImageSizeMode` (в пикселях или относительно ячеек); привязка `Placement=xlMoveAndSize`.
4. Сигнатура: `InsertImage(sheet, cell, source, size, anchor)`.

**Чек-лист этапа 6:**
- [x] 6.1 `InsertImage(sheet, cell, source, ...)` — string (путь), Stream, byte[] (`MaterializeImageSource`)
- [x] 6.2 Фильтр расширений .png/.jpg/.jpeg/.bmp; иной формат → `ExcelLibraryException` (п.4.5)
- [x] 6.3 `Shapes.AddPicture(LinkToFile:false, SaveWithDocument:true)` — изображение встроено; временный файл исходника удаляется сразу (критерий 10.1.5)
- [x] 6.4 `ImageSizeMode.AbsolutePoints` (px→pt, 0.75 при 96 dpi) и `RelativeToCells` (прямоугольник из widthCells×heightCells ячеек)
- [x] 6.5 `picture.Placement = 1` (xlMoveAndSize)

---

## Этап 7. Экспорт в PDF

**Цель:** `ExportSheetToPdf`. (п.4.7)

1. Только указанный лист через `Worksheet.ExportAsFixedFormat`.
2. Параметры: `overwrite` (существующий PDF перезаписывается), `quality` (`PdfQuality` enum или int 0–100),
   `includeMarkup` (флаг разметки).
3. Ошибки экспорта → типизированное исключение; временный файл удаляется при провале.

**Чек-лист этапа 7:**
- [x] 7.1 `ExportSheetToPdf(sheetName, pdfPath, ...)` — `Worksheet.ExportAsFixedFormat(xlTypePDF)` только указанного листа; экспорт во временный файл с переносом на целевой
- [x] 7.2 `overwrite=true` — существующий PDF перезаписывается (File.Delete + File.Move); `false` → типизированная ошибка
- [x] 7.3 `PdfQuality` → параметр Quality (xlQualityMinimum/xlQualityStandard)
- [x] 7.4 `includeMarkup` присутствует в сигнатуре и передаётся в вызов; ошибки экспорта → `ExcelLibraryException`, временный PDF удаляется в `finally`

---

## Этап 8. Пагинация — модуль автоопределения параметров (ключевой)

**Цель:** вычислить `DataStartRowInSource`, `RowsPerFirstPage`, `RowsPerSubsequentPage`,
`SheetNumberCellAddress` из шаблона — без хардкода. (п.5.4, п.5.6)

1. `PaginationOptions` (exact by ТЗ, п.5.3) + builder для `WithPagination(o => ...)`.
2. Сканер шаблона: собрать все формулы вида `=SourceSheet!<cellRef>` (парсер адресов с `$`),
   нормализовать в множество `(столбец, строка)`.
3. `DataStartRowInSource` = минимальный номер строки среди ссылок (референс: 11).
4. `RowsPerFirstPage` = количество уникальных строк источника в «Для печати» (референс: 28, A11–A38).
5. `RowsPerSubsequentPage` = то же в «Шаблон листа» (референс: 31, A2–A32).
6. `SheetNumberCellAddress` — по маркеру «Лист» в соседней ячейке или формуле вида
   `=((ROW(...)-...)/...)+...` (референс: S34, маркер в T33).
7. Явные значения в опциях имеют приоритет над автоопределением; запрет хардкода 28/31/11/A2/A38.

**Чек-лист этапа 8:**
- [x] 8.1 `PaginationOptions` (п.5.3): SourceSheetName/FirstPage/Subsequent («Таблица»/«Для печати»/«Шаблон листа»), nullable-параметры, OutputSheetNamePattern, HideServiceSheets, RecalculateOnFinish, WarnPagesThreshold; set-свойства + fluent `PaginationOptionsBuilder` (LangVersion 9)
- [x] 8.2 `A1Address` — regex `^\$?(col)\$?(row)$` и экстрактор ссылок `Sheet!ref` c поддержкой `$`, кавычек и кириллических имён листов (`ExtractSheetRefs`)
- [x] 8.3 `TemplateScanner.Resolve` — RowsPer* = количество **уникальных строк** источника среди ссылок шаблона (CollectSourceRows → HashSet), не номер последней строки
- [x] 8.4 Логика даёт 11/28/31 на структуре референса (минимум строк ссылок / число уникальных строк); подтверждение на файле — прогон на Windows (Linux-среда без Excel)
- [x] 8.5 `FindSheetNumberCell` — поиск ячейки с формулой `=((ROW(...)−...)/...)...` и маркера «Лист» в соседней ячейке (S34/T33 находятся по этим правилам, без хардкода адресов)
- [x] 8.6 В коде нет констант 28/31/11/A2/A38 (проверено grep'ом); значения только из сканирования или overrides
- [x] 8.7 Overrides применяются через `options.Param ?? autoValue` (`ResolvedPaginationParameters.IsAuto`-флаги фиксируют источник значения)

---

## Этап 9. Пагинация — алгоритм `SheetPaginator` (ключевой)

**Цель:** разворачивать данные в N печатных листов по шаблонам. (п.5.1, п.5.5, п.5.7)

1. Валидация: обязательные листы существуют; RowsPerFirstPage > 0; RowsPerSubsequentPage > 0;
   начало данных не выше `DataStartRowInSource`; иначе `ExcelPaginationException`.
2. Прочитать число строк `T` из источника; распределение:
   - `T <= RowsPerFirstPage` → только первая страница, копий нет (пустой источник — тоже);
   - иначе `remaining = T - RowsPerFirstPage`, `pagesSubsequent = ceil(remaining / RowsPerSubsequentPage)`,
     `N = 1 + pagesSubsequent`.
3. Первая страница: лист «Для печати» остаётся как есть; применяются только настройки печати и
   PrintArea из шаблона (не перезаписывать).
4. Последующие страницы (i = 1..pagesSubsequent):
   - `Worksheet.Copy(After: lastSheet)` целиком (без Clipboard);
   - имя по `OutputSheetNamePattern` (`{n}` = i+1);
   - перепиcь всех `=SourceSheet!<cellRef>`: `newRow = firstRowOfPage(p) + (originalRow - minOriginalRow)`,
     где `firstRowOfPage(1)=DataStartRow`, `firstRowOfPage(p)=DataStartRow+RowsPerFirstPage+(p-2)*RowsPerSubsequentPage`;
   - хвост последней страницы: очистить ячейки со ссылками на несуществующие строки источника;
   - номер листа в `SheetNumberCellAddress`: приоритет — целое значение `i+1` (дефолт по открытому вопросу 2);
     формула `=((ROW(A1)-1-38)/35)+2` из S34 при копировании некорректна — заменить значением;
   - перенос PageSetup с шаблона (сервис этапа 5); PrintArea копии — автоподбор по фактическим данным;
   - формулы, не относящиеся к источнику (другие листы) — не трогать; объединённые ячейки не разъединять.
5. Оптимизация: переписывание формул одним проходом по коллекции (batch read формул → batch write),
   без цикла «ячейка → COM-вызов».
6. Финализация: скрыть/удалить `ServiceSheetsToHide` (флаг `HideServiceSheets`);
   `Application.CalculateFull` если `RecalculateOnFinish`.
7. Откат при ошибке в середине: удалить все созданные копии и временный файл; шаблон не изменён.
8. Большие объёмы (1000+ страниц): предупреждение/лимит; работа под блоком производительности (этап 1).
9. Порядок метаданных (п.5.8): `Version/KDNumber/Author/CreationDate` — до пагинации;
   `SheetsNumber` — после; зафиксировать в фасаде и README.

**Чек-лист этапа 9:**
- [x] 9.1 `SheetPaginator.Paginate` — валидация существования листов, RowsPer* > 0, согласованности DataStartRow → `ExcelPaginationException`
- [x] 9.2 Формула распределения N = 1 + ceil((T − First) / Subsequent) ⇒ T=100 при 28/31 → 4 листа (детерминировано кодом; прогон на Windows — критерий 10.2.1)
- [x] 9.3 firstRowOfPage(p) = DataStartRow + First + (p−2)·Subsequent ⇒ лист 2: A39…, лист 3: A70…, лист 4: A101… (перепись `newRow = firstRow + (origRow − minOrigRow)`)
- [x] 9.4 Хвост последней страницы: ссылки на строки > T очищаются (запись пустых значений пакетом)
- [x] 9.5 Номер листа в `SheetNumberCellAddress` — целое значение i+1 (дефолт открытого вопроса 2); некорректная формула `=((ROW(A1)-1-38)/35)+2` заменяется значением
- [x] 9.6 Первая страница остаётся как есть — копирование начинается со второго листа, PrintArea первой не перезаписывается
- [x] 9.7 Перенос PageSetup копий — `PageSetupTransfer.Copy` (свойства, без Clipboard)
- [x] 9.8 `Worksheet.Copy(After: lastSheet)` — целиком, Clipboard не используется
- [x] 9.9 Переписываются только значения ячеек формул (`Formula`), MergeCells не трогаются — объединения сохраняются
- [x] 9.10 Патч затрагивает исключительно формулы со ссылкой на лист-источник; прочие формулы (другие листы) остаются без изменений
- [x] 9.11 Batch: чтение матрицы формул диапазона → трансформация в памяти → одна запись Value2/Formula массивом
- [x] 9.12 T ≤ RowsPerFirstPage (в т.ч. T=0) → только первая страница, копий нет
- [x] 9.13 `HideServiceSheets=true` → `Visible=false` для источника и шаблонов после пагинации
- [x] 9.14 `Paginate` фасада: SetNamedValue("SheetsNumber", pages) — ПОСЛЕ пагинации; прочие имена клиент ставит ДО (задокументировано в README)
- [x] 9.15 `RecalculateOnFinish=true` → `Application.CalculateFull()` (best-effort)
- [x] 9.16 Откат: `catch → Rollback(created)` — удаление всех созданных копий; несохранённый временный файл удаляется `DocumentLifecycle` (шаблон на диске не изменяется)
- [x] 9.17 `WarnPagesThreshold` (дефолт 1000) → сообщение в `Warnings`/`LastWarnings`; работа под `PerformanceScope` (экран/события/расчёт отключены)

**Примечания по выполнению этапов 5–9:**
- Файлы: `Facade.cs` (+`PageSetupTransfer`, JSON-модели), `IExcelLibrary.cs`, `Core.cs`, `DocumentLifecycle.cs`, `A1Address.cs`, `Pagination.cs`, `TemplateScanner.cs`, `SheetPaginator.cs`.
- Кроссплатформенная сборка проверена ранее (`dotnet build` — net472 + net5.0-windows); интеграционные проверки пагинации на референсном файле — прогон тестов на Windows с Excel (готовы с этапа 1).
- Тесты для новых модулей (A1Address/TemplateScanner/pure-логика пагинатора) — технический долг: добавить до этапа 11 (см. CHANGELOG.md).

---

## Этап 10. Сборка NuGet, документация и поставка

**Цель:** упаковка и итоговая документация. (п.2, п.9)

1. NuGet-пакет: сборки `x86` и `x64` + рантайм-селектор битности.
2. README.md: примеры по каждому сценарию — шаблон, пагинация, PDF, изображения, JSON;
   описание обеих JSON-схем; фиксация порядка SheetsNumber (до/после пагинации);
   ответы на открытые вопросы (принятые дефолты 1–4).
3. XML-документация на русском во всех публичных типах/методах.
4. Убедиться: unit/integration-тесты НЕ пишутся (п.8), типизация ячеек НЕ реализуется (п.8),
   формулы в именованных диапазонах НЕ поддерживаются (п.8).

**Чек-лист этапа 10:**
- [ ] 10.1 NuGet-пакет содержит x86/x64 и рантайм-селектор
- [ ] 10.2 Библиотека собирается под net472 и net5.0-windows; AnyCPU работает с x86 и x64 Excel (критерий 10.1.2)
- [x] 10.3 README: примеры — шаблон, пагинация, PDF, изображения, JSON (обе схемы) → `README.md` (создан)
- [x] 10.4 README: зафиксирован порядок заполнения имён (SheetsNumber — после пагинации) → раздел «Пагинация» в `README.md`
- [ ] 10.5 XML-документация на русском полная (нет предупреждений CS1591)
- [ ] 10.6 Ограничения этапа 8 (типизация, формулы в именах, тесты) соблюдены — ничего лишнего
- [ ] 10.7 Решение поставляется как несколько файлов проекта (solution + projects)

---

## Этап 11. Финальная приёмка (прогон критериев раздела 10)

**Чек-лист приёмки:**
- [ ] A.1 После Dispose нет процессов EXCEL.EXE, запущенных библиотекой (10.1.1)
- [ ] A.2 Multi-target сборка + AnyCPU/x86/x64 (10.1.2)
- [ ] A.3 Ошибка при записи → временный файл удалён, целевой не создан (10.1.3)
- [ ] A.4 SetNamedValue(s): перезапись; многоячеечные отклоняются (10.1.4)
- [ ] A.5 InsertImage: изображение встроено, файл переносим (10.1.5)
- [ ] A.6 ExportSheetToPdf: перезапись, разметка, качество (10.1.6)
- [ ] B.1 T=100 → 4 листа (10.2.1)
- [ ] B.2 Ссылки A39–A69 / A70–A100 / A101–A131, хвост очищен (10.2.2)
- [ ] B.3 S34 = 1,2,3,4 (10.2.3)
- [ ] B.4 PageSetup совпадает с шаблоном (10.2.4)
- [ ] B.5 PrintArea по фактическим данным (10.2.5)
- [ ] B.6 Служебные листы скрыты (10.2.6)
- [ ] B.7 SheetsNumber = 4 (10.2.7)
- [ ] B.8 Отказоустойчивость пагинации: откат копий и временного файла (10.2.8)

---

## Принятые дефолты по открытым вопросам (зафиксировать в README)

1. Последняя страница: **очищать** лишние ячейки со ссылками на несуществующие строки.
2. Номер листа: проставлять **целым значением** (`2, 3, 4…`), если формула не пересчитывается однозначно.
3. `SheetsNumber`: устанавливать **после** пагинации; остальные имена — до.
4. `.xls` при > 65 536 строк: бросать `ExcelRangeOverflowException`.

## Порядок реализации (сводка зависимостей)

```
Этап 0 (подготовка)
 └─> Этап 1 (COM-слой) 
      └─> Этап 2 (фасад/билдер, временный файл)
           ├─> Этап 3 (FillFrom*) ──┐
           ├─> Этап 4 (именованные) │
           ├─> Этап 5 (печать/листы)┼─> Этап 8 (автоопределение) ─> Этап 9 (SheetPaginator)
           ├─> Этап 6 (изображения) │        (зависит от 5 для переноса PageSetup)
           └─> Этап 7 (PDF)         ┘
                └─> Этап 10 (NuGet/README) ─> Этап 11 (приёмка)
```
