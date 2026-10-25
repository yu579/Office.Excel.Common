using System;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Параметры модуля пагинации (п.5.3 ТЗ). Свойства изменяемые (LangVersion 9: init-семантика
    /// достигается fluent-билдером <see cref="PaginationOptionsBuilder"/>); значения по умолчанию
    /// соответствуют референсному шаблону «338535-00-90-000 ПЭ3».
    /// Числовые параметры = null ⇒ автоопределение из шаблона (Этап 8, п.5.4); явные значения
    /// имеют приоритет над автоопределением (п.5.6 — запрет хардкода).
    /// </summary>
    public sealed class PaginationOptions
    {
        /// <summary>Имя листа-источника данных.</summary>
        public string SourceSheetName { get; set; } = "Таблица";

        /// <summary>Имя листа-шаблона первой страницы.</summary>
        public string FirstPageTemplateSheetName { get; set; } = "Для печати";

        /// <summary>Имя листа-шаблона последующих страниц.</summary>
        public string SubsequentTemplateSheetName { get; set; } = "Шаблон листа";

        /// <summary>Номер строки начала данных в источнике; null — автоопределение (минимум среди ссылок шаблона).</summary>
        public int? DataStartRowInSource { get; set; } = null;

        /// <summary>Число строк источника на первой странице; null — автоопределение (уникальные ссылки шаблона).</summary>
        public int? RowsPerFirstPage { get; set; } = null;

        /// <summary>Число строк источника на последующих страницах; null — автоопределение.</summary>
        public int? RowsPerSubsequentPage { get; set; } = null;

        /// <summary>Шаблон имён выходных листов; {n} заменяется на номер страницы (1-based).</summary>
        public string OutputSheetNamePattern { get; set; } = "Лист {n}";

        /// <summary>Скрывать ли служебные листы после пагинации.</summary>
        public bool HideServiceSheets { get; set; } = true;

        /// <summary>Служебные листы для скрытия (не удаляются — п.5.5.7 «скрыть/удалить» реализуется скрытием).</summary>
        public string[] ServiceSheetsToHide { get; set; } =
            new[] { "Таблица", "Info", "Шаблон листа" };

        /// <summary>Выполнить Application.CalculateFull после пагинации.</summary>
        public bool RecalculateOnFinish { get; set; } = true;

        /// <summary>Адрес ячейки номера листа на шаблонах; null — автоопределение по маркеру «Лист»/формуле ROW().</summary>
        public string? SheetNumberCellAddress { get; set; } = null;

        /// <summary>Предупреждение/лимит больших объёмов (п.5.7): максимум создаваемых страниц; 0 — без лимита.</summary>
        public int MaxPages { get; set; } = 1000;

        /// <summary>Порог, при котором выводится предупреждение о большом числе страниц.</summary>
        public int WarnPagesThreshold { get; set; } = 200;

        /// <summary>Копия опций (передача в пагинатор без риска внешней мутации).</summary>
        internal PaginationOptions Clone() => new PaginationOptions
        {
            SourceSheetName = SourceSheetName,
            FirstPageTemplateSheetName = FirstPageTemplateSheetName,
            SubsequentTemplateSheetName = SubsequentTemplateSheetName,
            DataStartRowInSource = DataStartRowInSource,
            RowsPerFirstPage = RowsPerFirstPage,
            RowsPerSubsequentPage = RowsPerSubsequentPage,
            OutputSheetNamePattern = OutputSheetNamePattern,
            HideServiceSheets = HideServiceSheets,
            ServiceSheetsToHide = (string[])ServiceSheetsToHide.Clone(),
            RecalculateOnFinish = RecalculateOnFinish,
            SheetNumberCellAddress = SheetNumberCellAddress,
            MaxPages = MaxPages,
            WarnPagesThreshold = WarnPagesThreshold,
        };
    }

    /// <summary>
    /// Результат автоопределения параметров пагинации из шаблона (Этап 8):
    /// значения, вычисленные сканером <see cref="TemplateScanner"/>, с флагом источника
    /// (авто/явный override) для каждой величины.
    /// </summary>
    public sealed class ResolvedPaginationParameters
    {
        /// <summary>Номер строки начала данных в источнике (DataStartRowInSource).</summary>
        public int DataStartRow { get; internal set; }

        /// <summary>Строк источника на первой странице.</summary>
        public int RowsPerFirstPage { get; internal set; }

        /// <summary>Строк источника на последующих страницах.</summary>
        public int RowsPerSubsequentPage { get; internal set; }

        /// <summary>Адрес ячейки номера листа на листах-шаблонах (например «S34») или null.</summary>
        public string? SheetNumberCellAddress { get; internal set; }

        /// <summary>Минимальный исходный номер строки ссылки шаблона первой страницы (для смещений, обычно = DataStartRow).</summary>
        public int MinOriginalRefRowFirstPage { get; internal set; }

        /// <summary>Минимальный исходный номер строки ссылки шаблона последующих страниц (обычно 2 — A2..A32).</summary>
        public int MinOriginalRefRowSubsequent { get; internal set; }

        /// <summary>Множество столбцов ссылок источника на первой странице (для пакетной перепи формул).</summary>
        internal int[] SourceColumnsFirstPage = Array.Empty<int>();

        /// <summary>Множество столбцов ссылок источника на последующих страницах.</summary>
        internal int[] SourceColumnsSubsequent = Array.Empty<int>();

        /// <summary>Диапазон «колонка:строки», который занимали ссылки на первой странице (для проверки PrintArea).</summary>
        internal int LastRefRowFirstPage;

        /// <summary>Максимальный номер строки ссылки на шаблоне последующих страниц.</summary>
        internal int LastRefRowSubsequent;

        /// <summary>Откуда взято значение DataStartRow: автоопределение или override из опций.</summary>
        public bool DataStartRowIsAuto { get; internal set; }

        /// <summary>RowsPerFirstPage определён автоматически (а не задан явно).</summary>
        public bool RowsPerFirstPageIsAuto { get; internal set; }

        /// <summary>RowsPerSubsequentPage определён автоматически.</summary>
        public bool RowsPerSubsequentPageIsAuto { get; internal set; }

        /// <summary>Ячейка номера листа найдена автоматически.</summary>
        public bool SheetNumberCellIsAuto { get; internal set; }
    }
}
