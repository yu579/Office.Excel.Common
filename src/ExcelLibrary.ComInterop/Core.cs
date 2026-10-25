using System;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Точка входа библиотеки: создаёт fluent-билдер конфигурации документа.
    /// </summary>
    /// <example>
    /// <code>
    /// using var excel = ExcelLibrary.Create()
    ///     .FromTemplate(@"C:\templates\form.xlsx")
    ///     .WithHiddenWorkFile(@"C:\out\result.xlsx")
    ///     .WithPagination(o =&gt; o.SourceSheet("Таблица"))
    ///     .Build();
    /// </code>
    /// </example>
    public static class ExcelLibrary
    {
        /// <summary>Начинает конфигурирование экземпляра библиотеки.</summary>
        /// <returns>Новый <see cref="ExcelLibraryBuilder"/>.</returns>
        public static ExcelLibraryBuilder Create() => new ExcelLibraryBuilder();
    }

    /// <summary>
    /// Режим сохранения целевого документа (определяется по расширению, п.4.4 ТЗ).
    /// </summary>
    public enum ExcelOutputFormat
    {
        /// <summary>Excel 97-2003 (.xls), лимит 65 536 строк на лист.</summary>
        Xls,

        /// <summary>Excel Workbook (.xlsx).</summary>
        Xlsx,

        /// <summary>Excel Macro-Enabled Workbook (.xlsm).</summary>
        Xlsm,
    }

    /// <summary>
    /// Способ задания размера изображения при вставке (п.4.5 ТЗ).
    /// </summary>
    public enum ImageSizeMode
    {
        /// <summary>Размер задаётся явно в пунктах (points; ~1 px = 0.75 pt при 96 dpi).</summary>
        AbsolutePoints,

        /// <summary>Изображение растягивается на заданное число ячеек (widthCells × heightCells).</summary>
        RelativeToCells,
    }

    /// <summary>
    /// Качество экспорта в PDF (передаётся в ExportAsFixedFormat, п.4.7 ТЗ).
    /// </summary>
    public enum PdfQuality
    {
        /// <summary>Минимальный размер файла (xlQualityMinimum = 0).</summary>
        Minimum = 0,

        /// <summary>Стандартное качество (xlQualityStandard = 1).</summary>
        Standard = 1,

        /// <summary>Высокое качество (для печати).</summary>
        High = 2,
    }

    /// <summary>
    /// Пресеты параметров пагинации из примера п.3 ТЗ: быстрый старт с дефолтами
    /// референсного шаблона («Таблица», «Для печати», «Шаблон листа»).
    /// Все числовые параметры остаются null — значения автоопределяются из шаблона (Этап 8).
    /// </summary>
    public sealed class PaginationOptionsBuilder
    {
        private readonly PaginationOptions _options = new PaginationOptions();

        /// <summary>Имя листа-источника данных.</summary>
        public PaginationOptionsBuilder SourceSheet(string name)
        {
            _options.SourceSheetName = name;
            return this;
        }

        /// <summary>Имя листа-шаблона первой страницы.</summary>
        public PaginationOptionsBuilder FirstPageTemplate(string name)
        {
            _options.FirstPageTemplateSheetName = name;
            return this;
        }

        /// <summary>Имя листа-шаблона последующих страниц.</summary>
        public PaginationOptionsBuilder SubsequentTemplate(string name)
        {
            _options.SubsequentTemplateSheetName = name;
            return this;
        }

        /// <summary>Скрывать ли служебные листы после пагинации.</summary>
        public PaginationOptionsBuilder HideServiceSheets(bool hide = true)
        {
            _options.HideServiceSheets = hide;
            return this;
        }

        /// <summary>Явное начало данных в источнике (override автоопределения).</summary>
        public PaginationOptionsBuilder DataStartRow(int? row)
        {
            _options.DataStartRowInSource = row;
            return this;
        }

        /// <summary>Явное число строк на первой странице (override автоопределения).</summary>
        public PaginationOptionsBuilder RowsPerFirstPage(int? rows)
        {
            _options.RowsPerFirstPage = rows;
            return this;
        }

        /// <summary>Явное число строк на последующих страницах (override автоопределения).</summary>
        public PaginationOptionsBuilder RowsPerSubsequentPage(int? rows)
        {
            _options.RowsPerSubsequentPage = rows;
            return this;
        }

        /// <summary>Шаблон имени выходных листов ({n} — номер страницы).</summary>
        public PaginationOptionsBuilder OutputSheetNamePattern(string pattern)
        {
            _options.OutputSheetNamePattern = pattern;
            return this;
        }

        /// <summary>Явный адрес ячейки номера листа (override автоопределения).</summary>
        public PaginationOptionsBuilder SheetNumberCellAddress(string? address)
        {
            _options.SheetNumberCellAddress = address;
            return this;
        }

        /// <summary>Пересчитывать книгу после пагинации (CalculateFull).</summary>
        public PaginationOptionsBuilder RecalculateOnFinish(bool recalc = true)
        {
            _options.RecalculateOnFinish = recalc;
            return this;
        }

        /// <summary>Служебные листы, скрываемые после пагинации.</summary>
        public PaginationOptionsBuilder ServiceSheets(params string[] names)
        {
            _options.ServiceSheetsToHide = names;
            return this;
        }

        /// <summary>Максимальное число страниц; 0 — без ограничения (п.5.7).</summary>
        public PaginationOptionsBuilder MaxPages(int maxPages)
        {
            _options.MaxPages = maxPages;
            return this;
        }

        /// <summary>Внутренний доступ к собранным опциям для билдера библиотеки.</summary>
        internal PaginationOptions Build() => _options;
    }

    /// <summary>
    /// Fluent-билдер конфигурации и сборки экземпляра <see cref="IExcelLibrary"/> (п.2 ТЗ).
    /// Поддерживает сценарии из примера п.3: шаблон/новая книга, скрытый рабочий файл, пагинация.
    /// Экземпляр пригоден для DI: <c>services.AddSingleton&lt;IExcelLibrary&gt;(sp =&gt; ExcelLibrary.Create()...Build())</c>.
    /// </summary>
    public sealed class ExcelLibraryBuilder
    {
        private string? _templatePath;
        private System.IO.Stream? _templateStream;
        private bool _useNewWorkbook;
        private string? _workFilePath;
        private PaginationOptions? _paginationOptions;
        private ExcelApplicationMode _mode = ExcelApplicationMode.CreateNew;
        private bool _checkBitness = true;

        /// <summary>Загрузить шаблон из файла по пути.</summary>
        /// <param name="path">Путь к .xls/.xlsx/.xlsm.</param>
        /// <exception cref="System.IO.FileNotFoundException">Файл не найден (проверяется в Build).</exception>
        public ExcelLibraryBuilder FromTemplate(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Путь к шаблону не задан.", nameof(path));
            _templatePath = path;
            _templateStream = null;
            _useNewWorkbook = false;
            return this;
        }

        /// <summary>Загрузить шаблон из потока (копия сохраняется во временный файл автоматически).</summary>
        /// <param name="stream">Поток с содержимым книги Excel.</param>
        public ExcelLibraryBuilder FromTemplate(System.IO.Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            _templateStream = stream;
            _templatePath = null;
            _useNewWorkbook = false;
            return this;
        }

        /// <summary>Создать новую пустую книгу без шаблона.</summary>
        public ExcelLibraryBuilder New()
        {
            _useNewWorkbook = true;
            _templatePath = null;
            _templateStream = null;
            return this;
        }

        /// <summary>
        /// Указать целевой путь документа: все операции выполняются в скрытом временном файле
        /// рядом с этим путём, а <see cref="IExcelLibrary.Save"/> выполняет SaveAs в видимый файл
        /// с форматом по расширению (п.4.4). При ошибке до успешного сохранения временный файл удаляется.
        /// </summary>
        /// <param name="savePath">Целевой путь (.xls/.xlsx/.xlsm).</param>
        public ExcelLibraryBuilder WithHiddenWorkFile(string savePath)
        {
            if (string.IsNullOrWhiteSpace(savePath)) throw new ArgumentException("Целевой путь не задан.", nameof(savePath));
            _workFilePath = savePath;
            return this;
        }

        /// <summary>Включить и настроить модуль пагинации (п.5.3).</summary>
        /// <param name="configure">Действие над пресетами параметров пагинации.</param>
        public ExcelLibraryBuilder WithPagination(Action<PaginationOptionsBuilder> configure)
        {
            var presets = new PaginationOptionsBuilder();
            configure?.Invoke(presets);
            _paginationOptions = presets.Build();
            return this;
        }

        /// <summary>Включить пагинацию с параметрами по умолчанию (прямая передача опций).</summary>
        /// <param name="options">Готовые параметры пагинации.</param>
        public ExcelLibraryBuilder WithPagination(PaginationOptions options)
        {
            _paginationOptions = options ?? throw new ArgumentNullException(nameof(options));
            return this;
        }

        /// <summary>Режим запуска Excel (по умолчанию — создание нового скрытого экземпляра).</summary>
        /// <param name="mode">Режим COM-приложения.</param>
        public ExcelLibraryBuilder WithApplicationMode(ExcelApplicationMode mode)
        {
            _mode = mode;
            return this;
        }

        /// <summary>Отключить проверку битности процесса и Excel.</summary>
        /// <param name="enabled">false — не проверять.</param>
        public ExcelLibraryBuilder WithoutBitnessCheck(bool enabled = false)
        {
            _checkBitness = enabled;
            return this;
        }

        /// <summary>
        /// Собирает экземпляр библиотеки: запускает Excel, открывает шаблон/новую книгу,
        /// подготавливает рабочий (скрытый временный) файл. Владелец результата — вызывающий код:
        /// <see cref="IDisposable.Dispose"/> освобождает и документ, и COM-приложение.
        /// </summary>
        /// <returns><see cref="IExcelLibrary"/>, реализующий IDisposable.</returns>
        /// <exception cref="InvalidOperationException">Источник документа не выбран (ни FromTemplate, ни New).</exception>
        /// <exception cref="ExcelTemplateException">Шаблон не открылся.</exception>
        public IExcelLibrary Build()
        {
            if (!_useNewWorkbook && _templatePath == null && _templateStream == null)
                throw new InvalidOperationException(
                    "Не выбран источник документа: вызовите FromTemplate(...) или New() перед Build().");

            var app = new ComExcelApplication(_mode, visible: false, checkBitness: _checkBitness);
            OwnedComExcelApplication? owned = null;
            try
            {
                owned = new OwnedComExcelApplication(app);
                var doc = new DocumentLifecycle(app, _workFilePath);

                if (_useNewWorkbook)
                {
                    doc.OpenNewWorkbook();
                }
                else if (_templatePath != null)
                {
                    doc.OpenFromTemplate(_templatePath!);
                }
                else
                {
                    doc.OpenFromStream(_templateStream!);
                }

                var facade = new ExcelLibraryFacade(doc, _paginationOptions);
                // владение приложением передаётся фасаду (цепочка Dispose)
                FacadeOwnership.Attach(facade, owned);
                owned = null;
                return facade;
            }
            catch
            {
                owned?.Dispose(); // освобождает и приложение, и книгу, и temp
                throw;
            }
        }

        /// <summary>Обёртка владения COM-приложением для корректной цепочки Dispose.</summary>
        internal sealed class OwnedComExcelApplication : IDisposable
        {
            /// <summary>COM-приложение.</summary>
            public ComExcelApplication App { get; }

            internal OwnedComExcelApplication(ComExcelApplication app) => App = app;

            public void Dispose() => App.Dispose();
        }
    }
}
