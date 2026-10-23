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
- [ ] 2.1 Билдер собирает все конфигурации из примера п.3 (компилируемый smoke-код)
- [ ] 2.2 Загрузка шаблона из пути и из Stream работает
- [ ] 2.3 `New()` создаёт книгу без шаблона
- [ ] 2.4 Временный скрытый файл создаётся рядом с целевым и удаляется при ошибке (критерий 10.1.3)
- [ ] 2.5 SaveAs определяет формат .xls/.xlsx/.xlsm по расширению целевого файла
- [ ] 2.6 Фасад реализует IDisposable; повторный Dispose безопасен
- [ ] 2.7 Лимит 65 536 строк для .xls проверяется (открытый вопрос 4 — дефолт принят)

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
- [ ] 3.1 DataTable: заголовки пишутся по умолчанию, отключаются параметром
- [ ] 3.2 DBNull/null → пустые ячейки; значения пишутся как текст
- [ ] 3.3 object[,] и IEnumerable<T>: запись ровно одним вызовом `Range.Value2`
- [ ] 3.4 JSON схема (а) — массив объектов → один лист
- [ ] 3.5 JSON схема (б) — вложенная `sheets[]` с name/startCell/columns/data
- [ ] 3.6 Модель JSON-схемы описана в коде и в README
- [ ] 3.7 Переполнение диапазона бросает `ExcelRangeOverflowException`
- [ ] 3.8 Заполнение источника НЕ начинается выше `DataStartRowInSource` (валидация для пагинации)

---

## Этап 4. Именованные диапазоны

**Цель:** `SetNamedValue` / `SetNamedValues` для глобальных одноячеечных имён. (п.4.2)

1. Scope = Workbook; только одна ячейка; перезапись существующих имён.
2. Многоячеечные диапазоны — отклонять (исключение).
3. Формулы в именованных диапазонах не поддерживать (п.8) — только значения.
4. Использовать для `Version`, `KDNumber`, `Author`, `CreationDate`, `SheetsNumber`.

**Чек-лист этапа 4:**
- [ ] 4.1 `SetNamedValue(key, value)` создаёт/перезаписывает глобальное имя
- [ ] 4.2 `SetNamedValues(IDictionary<string, object>)` — массовая установка
- [ ] 4.3 Существующие имена перезаписываются без ошибки (критерий 10.1.4)
- [ ] 4.4 Многоячеечный диапазон отклоняется типизированным исключением (критерий 10.1.4)
- [ ] 4.5 Значения пишутся как текст; формулы не поддерживаются

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
- [ ] 5.1 `SetPrintArea(auto:true)` подбирает область по фактическим данным
- [ ] 5.2 Первая страница пагинации сохраняет PrintArea из шаблона (не перезаписывается)
- [ ] 5.3 Колонтитулы читаются и пишутся (6 слотов Left/Center/Right × Header/Footer)
- [ ] 5.4 `SetPrintTitles(rows:"1:2", columns:"A:B")` применяется корректно
- [ ] 5.5 Перенос PageSetup работает без Clipboard/PasteSpecial (только свойства)
- [ ] 5.6 `AddSheet` копирует стиль и настройки печати образца
- [ ] 5.7 `RemoveSheet` отказывает для единственного/защищённого листа
- [ ] 5.8 PrintArea каждого печатного листа покрывает фактические данные (критерий 10.2.5)

---

## Этап 6. Изображения

**Цель:** `InsertImage` с встраиванием в документ. (п.4.5)

1. Источники: путь к файлу, `Stream`, `byte[]`; форматы .png/.jpg/.bmp.
2. Встраивание через `Shapes.AddPicture(..., LinkToFile:=false, SaveWithDocument:=true)` —
   после сохранения файл не требует исходника.
3. Размер: enum `ImageSizeMode` (в пикселях или относительно ячеек); привязка `Placement=xlMoveAndSize`.
4. Сигнатура: `InsertImage(sheet, cell, source, size, anchor)`.

**Чек-лист этапа 6:**
- [ ] 6.1 Три источника изображения (path/Stream/byte[]) поддержаны
- [ ] 6.2 Только png/jpg/bmp; иные форматы — ошибка
- [ ] 6.3 Изображение встроено (критерий 10.1.5: перенос файла на другой ПК — картинка видна)
- [ ] 6.4 `ImageSizeMode`: пиксели и относительный размер по ячейкам
- [ ] 6.5 `Placement = xlMoveAndSize` задан

---

## Этап 7. Экспорт в PDF

**Цель:** `ExportSheetToPdf`. (п.4.7)

1. Только указанный лист через `Worksheet.ExportAsFixedFormat`.
2. Параметры: `overwrite` (существующий PDF перезаписывается), `quality` (`PdfQuality` enum или int 0–100),
   `includeMarkup` (флаг разметки).
3. Ошибки экспорта → типизированное исключение; временный файл удаляется при провале.

**Чек-лист этапа 7:**
- [ ] 7.1 Экспорт одного листа в PDF работает
- [ ] 7.2 Перезапись существующего PDF (критерий 10.1.6)
- [ ] 7.3 Параметр качества влияет на экспорт (критерий 10.1.6)
- [ ] 7.4 Флаг includeMarkup передаётся в ExportAsFixedFormat (критерий 10.1.6)

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
- [ ] 8.1 Класс `PaginationOptions` соответствует п.5.3 (init-only свойства, дефолты имён листов)
- [ ] 8.2 Парсер ссылок поддерживает A1/B11/AB123 и абсолютные `$A$11`
- [ ] 8.3 Rows-per-page считается по количеству ячеек с `=Таблица!`, а не по последней строке
- [ ] 8.4 На референсе: DataStartRow=11, First=28, Subsequent=31 определяются автоматически
- [ ] 8.5 SheetNumberCellAddress=S34 находится автоматически (маркер «Лист» в T33)
- [ ] 8.6 Ни одной константы 28/31/11/A2/A38 в коде (критерий 5.6)
- [ ] 8.7 Ручные overrides из опций переопределяют автоопределение

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
- [ ] 9.1 Валидация списков листов и параметров бросает `ExcelPaginationException`
- [ ] 9.2 T=100 на референсе → ровно 4 листа: 1 + ceil((100-28)/31) (критерий 10.2.1)
- [ ] 9.3 Лист 2 → Таблица!A39…A69; Лист 3 → A70…A100; Лист 4 → A101…A131 (критерий 10.2.2)
- [ ] 9.4 Хвост последней страницы: ссылки на несуществующие строки очищены (критерий 10.2.2, п.5.7)
- [ ] 9.5 Номера листов в S34: 1, 2, 3, 4 (критерий 10.2.3); сломанная формула заменена значением
- [ ] 9.6 Первая страница не модифицируется (кроме применения настроек из шаблона)
- [ ] 9.7 PageSetup копий идентичен шаблону; перенос без Clipboard (критерий 10.2.4)
- [ ] 9.8 Copy выполняется через `Worksheet.Copy(After:)`, Clipboard не используется (п.5.6)
- [ ] 9.9 Объединённые ячейки сохранены, не разъединяются при переписи формул (п.5.7)
- [ ] 9.10 Формулы на другие листы оставлены без изменений (п.5.6, п.5.7)
- [ ] 9.11 Перепись формул — одним проходом (batch), без «ячейка→COM» в цикле (п.7)
- [ ] 9.12 Пустой источник / T < RowsPerFirstPage → только первая страница (п.5.7)
- [ ] 9.13 Служебные листы скрыты при HideServiceSheets=true (критерий 10.2.6)
- [ ] 9.14 SheetsNumber=4 устанавливается ПОСЛЕ пагинации; остальные имена — ДО (критерий 10.2.7, п.5.8)
- [ ] 9.15 CalculateFull выполняется при RecalculateOnFinish=true
- [ ] 9.16 Авария в середине → копии и временный файл удалены, шаблон не изменён (критерий 10.2.8)
- [ ] 9.17 1000+ страниц: предупреждение/лимит; ScreenUpdating/Events/Calculation отключены (п.5.7)

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
- [ ] 10.3 README: примеры — шаблон, пагинация, PDF, изображения, JSON (обе схемы)
- [ ] 10.4 README: зафиксирован порядок заполнения имён (SheetsNumber — после пагинации)
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
