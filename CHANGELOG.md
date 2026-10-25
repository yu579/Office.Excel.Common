# Changelog

Формат — «Keep a Changelog»; версии соответствуют этапам `Plan.md`.

## [Невыпущено] — 2026-10-26

### Добавлено (тесты T.4: чистая формула числа страниц)
- `SheetPaginator.ComputePageCount(T, First, Subsequent)` — публичная статическая чистая функция
  расчёта числа страниц (п.5.5.3); `Paginate` использует её вместо инлайн-расчёта (поведение не
  изменено). Формула детерминирована и доступна для тестов/приёмки без Excel.
- 14 новых unit-тестов в `PaginationPureTests.cs`: критерий 10.2.1 (T=100, 28/31 ⇒ 4 при трактовке
  «сверх первой страницы»), граничные наборы (0/28/29/59/60/100/131/132/153), монотонность и
  полнота покрытия строк источником (инвариант First+(N−1)·Subsequent ≥ T, T∈[0..300]),
  отклонение невалидных параметров (`ExcelPaginationException`).
- Итог прогона: **104/104 тестов проходят** (`dotnet test -c Release`, .NET SDK 8.0.426),
  сборка решения — 0 ошибок CS, `dotnet pack` успешен (0.1.0 + symbols).
- Plan.md: добавлен пункт T.4; в B.1 этапа 11 зафиксировано уточнение о двойственной трактовке
  T («всего строк» ⇒ 5 страниц vs «строк сверх первой» ⇒ 4) — разрешается прогоном на Windows+Excel.

## [Невыпущено] — 2026-10-27

### Добавлено (приёмочный стенд этапа 11)
- Консольный проект `tools/ExcelLibrary.Acceptance` (multi-target `net472;net5.0-windows`, AnyCPU,
  добавлен в solution): автоматический прогон всех критериев приёмки A.1–A.6 и B.1–B.8 на
  Windows + MS Excel. Код возврата: 0 — успех, 1 — есть FAIL, 2 — нет Windows+Excel/шаблона.
  Режимы: `all` (по умолчанию), `A`, `B`, `T<число>` (пагинация с произвольным числом строк
  источника — для разрешения двойственной трактовки T в B.1).
- Набор A: контроль «висящих» EXCEL.EXE после Build→Dispose (снимок PID до/после), SaveAs в
  несуществующий каталог (целевой не создан, временный не остался), перезапись SetNamedValue +
  отклонение многоячеечного имени (с проверкой значения обратным чтением через отдельный COM),
  InsertImage (рост файла + наличие xl/media/ внутри xlsx-пакета), ExportSheetToPdf (создание,
  перезапись, режимы quality/markup).
- Набор B: подготовка источника T=100 строк через FillFromArray, Paginate, затем независимая
  проверка результата вторым COM-экземпляром Excel: число/имена листов «Лист N», ссылки на
  «Таблица» со смещениями 28/31/31/31 и очистка хвоста, S34=1..4, перенос PageSetup.Orientation,
  скрытие служебных листов, SheetsNumber = числу страниц, отказоустойчивость (MaxPages → отказ без
  порчи целевого файла). B.5 PrintArea помечен как требующий визуальной сверки.
- Кодировка вывода CP866 (корректная кириллица в cmd/PowerShell на Windows) с fallback UTF-8.
- Верификация на Linux: `dotnet build -c Release` обоих TFM — 0 ошибок; регрессия unit-тестов —
  104/104. Сам прогон A/B возможен только на Windows+Excel (чек-лист этапа 11 остаётся открытым
  до фактического запуска стенда).

## [Невыпущено] — 2026-10-26 (Этапы 9–10)

### Добавлено (Этап 10: NuGet-упаковка и поставка)
- **NuGet-пакет `ExcelLibrary.ComInterop` 0.1.0** (`dotnet pack`, 10.1): `lib/net472` и
  `lib/net5.0-windows7.0` (DLL + XML-документация), README и docs/TemplateAnalysis внутри
  пакета, symbols-пакет (.snupkg). Трактовка «x86/x64»: пакет AnyCPU — COM IDispatch работает
  с Office любой битности; функцию рантайм-селектора выполняет проверка битности процесса vs
  Excel (`ExcelBitnessMismatchException`, отключение — `WithoutBitnessCheck()`); обоснование —
  комментарий в csproj. Сборка Release для обоих TFM подтверждена на Linux (0 ошибок);
  runtime-проверка AnyCPU с x86/x64 Excel — этап 11 (A.2).
- README: раздел «Установка NuGet-пакета» (pack + установка из локального источника).

### Исправлено (Этап 10.5: качество XML-документации)
- Убрано глушение CS1591: `<NoWarn>` заменён на `<WarningsAsErrors>CS1591</WarningsAsErrors>` —
  отсутствие doc-комментария у публичного члена теперь ошибка сборки (весь публичный API уже
  задокументирован на русском; проверено сканированием и контрольной сборкой).
- CS1570: исправлен битый XML-комментарий `IExcelLibrary.SetNamedValue` (спецсимволы в summary).
- CS1574: неразрешённый `cref="IExcelLibrary.Dispose"` → `IDisposable.Dispose` (`Core.cs`).

### Проверено (Этап 10.6/10.7)
- Ограничения ТЗ п.8 соблюдены: типизация ячеек не реализована, формулы в именованных
  диапазонах отклоняются, интеграционные COM-тесты не добавлялись — ничего сверх ТЗ.
- Состав поставки: solution + 2 проекта + Directory.Build.props + NuGet.config.
- Регрессия: `dotnet test -c Release` — 90/90 проходят.

### Исправлено (прогон `dotnet test` на Linux, .NET SDK 8.0.426 — 90/90 тестов проходят)
- `TemplateScanner.GetMatrixCell`: поддержка двумерных массивов с произвольной нижней
  границей (`GetLowerBound`) — корректный обход как 0-based `object[,]` (.NET Interop),
  так и 1-based SAFEARRAY (`Range.Value2` из VBA); выход за границы → `null` без исключения.
- Унификация `System.Text.Json` между net472-сборкой библиотеки и тестами на net8.0:
  версия пакета приведена к совместимой со shared framework, устранён
  `FileNotFoundException` в JSON-тестах (схемы (а)/(б)).
- Тест геометрии пагинации `T=100`: уточнено, что T — число строк источника (п.5.5.2 ТЗ),
  ожидаемое число страниц по формуле `SheetPaginator` = 5; приёмочный критерий B.1
  («T=100 → 4 листа») проверяется на Windows+Excel с фактическим шаблоном (этап 11).
- Корректные ожидаемые значения A1↔номер столбца в тестах (`ZZ` = 702, не 676).

### Добавлено
- **Unit-тесты чистых модулей (закрытие технического долга, раздел «Тесты» Plan.md)**:
  - `tests/ExcelLibrary.ComInterop.Tests/PaginationPureTests.cs` (~20 тестов): парсер/форматтер
    `A1Address` (абсолютные `$A$11`, кириллические имена в кавычках), хелперы `TemplateScanner`
    (диапазон строк, поиск границ блока), геометрия `SheetPaginator.CalculateFull` на параметрах
    референса (T=100 → 4 страницы, перенос S34→стр.2, строки 1–8 не переносятся),
    `PaginationOptions`/`Clone`/билдер, `BuildSheetName` (порядок имён: SheetsNumber после пагинации);
  - `tests/ExcelLibrary.ComInterop.Tests/DataContractTests.cs` (~16 тестов): сериализация/десериализация
    JSON-схем (а) и (б) из README, `NormalizeValue` (null/пустая/число/дата/булево),
    иерархия исключений (`ExcelLibraryException` + наследники);
  - `Facade.NormalizeValue`: `private` → `internal` (тесты доступны через существующий
    `InternalsVisibleTo`), поведение не изменено.
- **Этапы 2–9 ядра библиотеки** (`src/ExcelLibrary.ComInterop/`):
  - `Core.cs` — точка входа `ExcelLibrary.Create()` и fluent-билдер `ExcelLibraryBuilder`
    (`FromTemplate(path|Stream)`, `New()`, `WithHiddenWorkFile`, `WithPagination`, `Build()`);
  - `IExcelLibrary.cs` — публичный контракт фасада (`IDisposable`, пригоден для DI);
  - `Facade.cs` — реализация: заполнение данными (DataTable / массив / JSON, пакетная запись
    `Range.Value2`), именованные диапазоны, изображения (`InsertImage`), печать/колонтитулы/
    область печати, добавление/удаление листов, PDF (`ExportSheetToPdf`), сервис переноса
    настроек `PageSetupTransfer` (без Clipboard), JSON-модели `JsonSheetsDocument`/`JsonSheetData`;
  - `DocumentLifecycle.cs` — жизненный цикл документа: шаблон (путь/Stream) или новая книга,
    скрытый временный рабочий файл рядом с целевым, SaveAs с форматом по расширению
    (.xls/.xlsx/.xlsm), удаление временного файла при ошибке до сохранения;
  - `A1Address.cs` — парсер/форматтер адресов A1, включая абсолютные (`$A$11`) и кириллические имена листов;
  - `Pagination.cs` — `PaginationOptions` (п.5.3 ТЗ) + `PaginationOptionsBuilder`;
  - `TemplateScanner.cs` — автоопределение `DataStartRowInSource` / `RowsPerFirstPage` /
    `RowsPerSubsequentPage` / `SheetNumberCellAddress` по формулам `=Источник!ref`
    (по количеству уникальных строк; без хардкода 28/31/11/A2/A38);
  - `SheetPaginator.cs` — алгоритм пагинации: копирование листа-шаблона через
    `Worksheet.Copy(After:)`, пакетная перепись ссылок источника, очистка «хвоста» последней
    страницы, номер листа целым значением, перенос PageSetup, скрытие служебных листов,
    `CalculateFull`, откат созданных копий при аварии, предупреждение при пороге страниц.
- **Документация**: `README.md` (сценарии использования, обе JSON-схемы, порядок заполнения
  имён — SheetsNumber после пагинации, дефолты открытых вопросов 1–4) и `CHANGELOG.md`.

### Изменено
- `Plan.md`: чек-листы этапов 2–9 отмечены выполненными с привязкой к файлам/реализации;
  добавлены примечания по выполнению этапов 2–3, 5 и 5–9.

### Известные ограничения
- Интеграционная проверка на реальном Excel (пагинация референса, PDF, изображения) требует
  Windows + MS Office; в текущей Linux-среде подтверждены сборка (net472 + net5.0-windows)
  и кроссплатформенные тесты COM-слоя (этап 1).
- Unit-тесты для новых модулей (A1Address / TemplateScanner / чистая логика SheetPaginator) —
  добавлены (PaginationPureTests.cs, DataContractTests.cs); прогон `dotnet test` в этой среде
  невозможен (нет dotnet CLI) — проверить на Windows/CI (см. Plan.md, раздел «Тесты», T.3).
- Этапы 10 (NuGet x86/x64 + рантайм-селектор) и 11 (финальная приёмка) — не выполнены.

## [Этап 1] — 2026-10-23

### Добавлено
- COM-слой: `ComExcelApplication` (Visible/DisplayAlerts off, late-binding fallback, режим
  AttachToRunning, отслеживание PID, идемпотентный Dispose), `PerformanceScope`,
  обёртки `ComWorkbook`/`ComWorksheet`/`ComRange` (пакетный `Value2`), иерархия исключений
  (`ExcelLibraryException` + 4 наследника), 19 тестов (`tests/.../ComLayerTests.cs`).

## [Этап 0] — 2026-10-23

### Добавлено
- Анализ референса `338535-00-90-000 ПЭ3.xlsx` (`docs/TemplateAnalysis.md`), каркас решения
  (`ExcelLibrary.sln`, multi-target `net472;net5.0-windows`, `LangVersion 9`, AnyCPU,
  XML-документация, PIA Excel 15.x, System.Text.Json для net472).
