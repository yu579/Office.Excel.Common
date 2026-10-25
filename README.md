# ExcelLibrary.ComInterop

Библиотека для работы с Microsoft Excel через **COM Interop** (C#, `LangVersion 9`).
Предназначена для заполнения корпоративных печатных форм по шаблону (референс —
`338535-00-90-000 ПЭ3.xlsx`) с **автоматической пагинацией** на N печатных листов,
экспортом в PDF и встраиванием изображений.

- Целевые фреймворки: `net472` и `net5.0-windows` (multi-target, AnyCPU).
- Требования: установленный Microsoft Excel (x86 или x64) на машине выполнения; ОС — Windows.
- Зависимости: PIA `Microsoft.Office.Interop.Excel` 15.x (с fallback на late binding), `System.Text.Json`.
- Анализ эталонного шаблона: [`docs/TemplateAnalysis.md`](docs/TemplateAnalysis.md). План работ: [`Plan.md`](Plan.md). История изменений: [`CHANGELOG.md`](CHANGELOG.md).

## Быстрый старт

```csharp
using ExcelLibrary.ComInterop;

using var excel = ExcelLibrary.Create()
    .FromTemplate(@"C:\templates\338535-00-90-000 ПЭ3.xlsx") // путь или Stream
    .WithHiddenWorkFile(@"C:\out\result.xlsx")               // целевой «видимый» файл
    .WithPagination(o => o
        .SourceSheet("Таблица")
        .FirstPageTemplate("Для печати")
        .SubsequentTemplate("Шаблон листа"))
    .Build();

// 1. Метаданные — ДО пагинации (п.5.8 ТЗ)
excel.SetNamedValues(new Dictionary<string, object>
{
    ["Version"]       = "1.0",
    ["KDNumber"]      = "338535-00-90-000",
    ["Author"]        = "Иванов И.И.",
    ["CreationDate"]  = DateTime.Now.ToString("dd.MM.yyyy"),
});

// 2. Данные в лист-источник
excel.FillFromDataTable("Таблица", "A11", myDataTable);

// 3. Пагинация: SheetsNumber устанавливается АВТОМАТИЧЕСКИ после разворачивания
int pages = excel.Paginate();   // 1 + ceil((T - RowsPerFirstPage) / RowsPerSubsequentPage)

// 4. Сохранение (формат определяется по расширению: .xls / .xlsx / .xlsm)
excel.Save();
```

Все операции выполняются в **скрытом временном рабочем файле** рядом с целевым; при любой
ошибке до успешного `Save()` временный файл удаляется, исходный шаблон не изменяется.
`Dispose()` идемпотентен и гарантированно освобождает COM-ресурсы (чужие экземпляры Excel не затрагиваются).

## Fluent-билдер (`ExcelLibraryBuilder`)

| Метод | Назначение |
|---|---|
| `FromTemplate(string path)` | Открыть шаблон по пути (.xls/.xlsx/.xlsm) |
| `FromTemplate(Stream stream)` | Открыть шаблон из потока (копируется во временный файл) |
| `New()` | Создать новую пустую книгу без шаблона |
| `WithHiddenWorkFile(string savePath)` | Целевой путь; работа — в скрытом временном файле рядом |
| `WithPagination(Action<PaginationOptionsBuilder>)` | Включить и настроить пагинацию |
| `Build()` | Собрать `IExcelLibrary` (реализует `IDisposable`, пригоден для DI) |

## API фасада `IExcelLibrary`

### Заполнение данными (одна пакетная запись `Range.Value2` на блок)
- `FillFromDataTable(sheetName, startCell, table, writeHeaders = true)` — `DBNull`/`null` → пустая ячейка, значения пишутся как текст.
- `FillFromArray(sheetName, startCell, object[,])` — двумерный массив.
- `FillFromArray<T>(sheetName, startCell, IEnumerable<T>, writeHeaders = false)` — свойства `T` → столбцы.
- `FillFromJson(sheetName, json)` — обе схемы, см. ниже.

Диапазон задаётся **только верхней левой ячейкой**; фактический прямоугольник вычисляется
и проверяется на переполнение: `.xls` > 65 536 строк или `.xlsx/.xlsm` > 1 048 576 строк
(> 16 384 столбцов) → `ExcelRangeOverflowException`.

### Именованные диапазоны (глобальные, workbook-scope, одноячеечные)
- `SetNamedValue(key, value)` / `SetNamedValues(IDictionary<string, object>)` — создание и **перезапись** имён (`Version`, `KDNumber`, `Author`, `CreationDate`, `SheetsNumber`).
- Многоячеечная ссылка в имени → `ExcelTemplateException`. Формулы в именованных диапазонах не поддерживаются (только значения).

### JSON-схемы
Схема (а) — массив объектов → один лист (`sheetName` и `startCell` берутся из параметров вызова):
```json
[
  { "Наименование": "Стопор", "Масса": "1,5" },
  { "Наименование": "Втулка", "Масса": "0,3" }
]
```

Схема (б) — несколько листов:
```json
{
  "sheets": [
    {
      "name": "Таблица",
      "startCell": "A11",
      "columns": ["Наименование", "Масса"],
      "data": [
        { "Наименование": "Стопор", "Масса": "1,5" }
      ]
    }
  ]
}
```
Модели в коде: `JsonSheetsDocument` / `JsonSheetData` (`Facade.cs`). Для схемы (б) параметр `sheetName` игнорируется (лист берётся из `name`).

### Изображения
- `InsertImage(sheetName, cell, source, widthPx, heightPx, sizeMode, widthCells, heightCells)` — источники: путь, `Stream`, `byte[]`; форматы `.png/.jpg/.bmp`.
- Встраивание через `Shapes.AddPicture(LinkToFile:=false, SaveWithDocument:=true)` — после сохранения файл не требует исходника; `Placement = xlMoveAndSize`.
- `ImageSizeMode.AbsolutePoints` (размер в px, конвертация в pt) или `RelativeToCells` (размер по прямоугольнику ячеек).

### Печать и листы
- `SetPrintArea(sheetName, auto: true, printArea = null)` — автоподбор по UsedRange.
- `SetHeaderFooter(sheetName, header, footer)` / `GetHeaderFooter(sheetName)` — 6 слотов Left/Center/Right × Header/Footer (`HeaderFooter`).
- `SetPrintTitles(sheetName, rows, columns)` — повторяющиеся строки/столбцы («1:2», «A:B»).
- `AddSheet(name, after?, modelSheet?)` — новый лист с переносом настроек образца (без Clipboard).
- `RemoveSheet(name)` — отказ для единственного/защищённого листа (`ExcelTemplateException`).

### Экспорт в PDF
- `ExportSheetToPdf(sheetName, pdfPath, overwrite = true, quality = PdfQuality.Standard, includeMarkup = false)` — экспорт только указанного листа через `ExportAsFixedFormat`; ошибки → типизированное исключение.

### Пагинация
- `Paginate()` (опции из `WithPagination`) / `Paginate(PaginationOptions options)` → число страниц.
- `LastWarnings` — предупреждения последней пагинации (например, порог `WarnPagesThreshold`, по умолчанию 1000 страниц).

**Порядок заполнения имён (зафиксирован, п.5.8):** `Version` / `KDNumber` / `Author` / `CreationDate` — **до** пагинации; `SheetsNumber` — **после** (библиотека проставляет его автоматически в `Paginate`).

## Автоопределение параметров пагинации (без хардкода)

`TemplateScanner` вычисляет параметры по формулам вида `=Таблица!A11` на листах-шаблонах:

| Параметр | Правило | Референс |
|---|---|---|
| `DataStartRowInSource` | минимальная строка среди ссылок шаблона первой страницы | 11 |
| `RowsPerFirstPage` | количество **уникальных** строк источника в «Для печати» | 28 (A11–A38) |
| `RowsPerSubsequentPage` | то же в «Шаблон листа» | 31 (A2–A32) |
| `SheetNumberCellAddress` | ячейка с маркером «Лист» в соседней ячейке либо с формулой `=((ROW(...)-...)/...)+...` | S34 |

Явные значения в `PaginationOptions` имеют приоритет над автоопределением. Констант 28/31/11/A2/A38 в коде нет.

Алгоритм `SheetPaginator`: первая страница не модифицируется; последующие — `Worksheet.Copy(After:)`
(без Clipboard), пакетная перепись ссылок источника, очистка «хвоста» последней страницы, номер листа
целым значением, перенос PageSetup свойствами, скрытие служебных листов, откат копий при аварии.

## Исключения

Иерархия (все — наследники `ExcelLibraryException`):

| Исключение | Когда возникает |
|---|---|
| `ExcelRangeOverflowException` | данные выходят за пределы листа/страницы (`RequestedCount`/`MaximumCount`) |
| `ExcelTemplateException` | проблемы шаблона: отсутствующие листы/имена, многоячеечное имя, удаление единственного/защищённого листа (`ElementName`) |
| `ExcelBitnessMismatchException` | разрядность процесса ≠ разрядность установленного Excel (сообщение содержит инструкцию) |
| `ExcelPaginationException` | ошибки конфигурации/выполнения пагинации |

## Принятые дефолты по открытым вопросам (см. Plan.md)

1. Источник данных — лист-«Таблица»; строки выше `DataStartRowInSource` считаются шапкой.
2. Номер печатного листа — **целое значение** i+1 (не формула).
3. `SheetsNumber` проставляется библиотекой **после** пагинации.
4. Лимит `.xls` — 65 536 строк, превышение → `ExcelRangeOverflowException`.

## Ограничения (по ТЗ, п.8)

- Типизация ячеек не реализуется — всё пишется как текст.
- Формулы в именованных диапазонах не поддерживаются.
- Unit/integration-тесты на COM-поведение требуют Windows с Excel; в Linux-среде доступны сборка и кроссплатформенные тесты контрактного слоя.

## Сборка и тесты

```powershell
dotnet build ExcelLibrary.sln          # net472 + net5.0-windows
dotnet test  tests/ExcelLibrary.ComInterop.Tests   # Windows-only интеграционные пропускаются graceful
```
