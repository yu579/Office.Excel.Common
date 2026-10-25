# Changelog

Формат — «Keep a Changelog»; версии соответствуют этапам `Plan.md`.

## [Невыпущено] — 2026-10-25

### Добавлено
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
- Unit-тесты для новых модулей (A1Address / TemplateScanner / чистая логика SheetPaginator)
  — технический долг, добавить до этапа 11.
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
