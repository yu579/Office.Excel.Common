using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Публичный фасад библиотеки (п.3 ТЗ): все методы работы с документом — заполнение данными,
    /// именованные диапазоны, изображения, печать, PDF, листы, пагинация, сохранение.
    /// Экземпляр реализует <see cref="IDisposable"/> (идемпотентный Dispose), пригоден для DI.
    /// </summary>
    public interface IExcelLibrary : IDisposable
    {
        // —— Заполнение данными (Этап 3, п.4.1) ——

        /// <summary>Заполнить лист из DataTable, начиная с верхней левой ячейки.</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="startCell">Адрес первой ячейки (A1).</param>
        /// <param name="table">Источник данных; DBNull/null → пустая ячейка.</param>
        /// <param name="writeHeaders">Писать строку заголовков (по умолчанию true).</param>
        void FillFromDataTable(string sheetName, string startCell, System.Data.DataTable table, bool writeHeaders = true);

        /// <summary>Заполнить лист из двумерного массива (одна запись Range.Value2).</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="startCell">Адрес первой ячейки.</param>
        /// <param name="data">Массив значений [строки, столбцы].</param>
        void FillFromArray(string sheetName, string startCell, object[,] data);

        /// <summary>Заполнить лист последовательностью объектов (свойства T → столбцы).</summary>
        /// <typeparam name="T">Тип элементов.</typeparam>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="startCell">Адрес первой ячейки.</param>
        /// <param name="items">Элементы.</param>
        /// <param name="writeHeaders">Писать заголовки из имён свойств.</param>
        void FillFromArray<T>(string sheetName, string startCell, IEnumerable<T> items, bool writeHeaders = false);

        /// <summary>Заполнить лист из JSON (схемы (а) и (б), п.4.1). Для схемы (б) параметр sheetName игнорируется.</summary>
        /// <param name="sheetName">Лист по умолчанию для схемы (а).</param>
        /// <param name="json">Текст JSON.</param>
        void FillFromJson(string sheetName, string json);

        // —— Именованные диапазоны (Этап 4, п.4.2) ——

        /// <summary>Создать/перезаписать глобальное одноячеечное имя.</summary>
        /// <param name="key">Имя (Version, KDNumber, …).</param>
        /// <param name="value">Значение (пишется как текст).</summary>
        void SetNamedValue(string key, object value);

        /// <summary>Массовая установка глобальных имён.</summary>
        /// <param name="values">Пары имя→значение.</param>
        void SetNamedValues(IDictionary<string, object> values);

        // —— Изображения (Этап 6, п.4.5) ——

        /// <summary>Вставить изображение в ячейку (встраивается в документ).</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="cell">Ячейка-якорь (верхний левый угол).</param>
        /// <param name="source">Путь к файлу, Stream или byte[] (png/jpg/bmp).</param>
        /// <param name="widthPx">Ширина в пикселях (для AbsolutePoints).</param>
        /// <param name="heightPx">Высота в пикселях (для AbsolutePoints).</param>
        /// <param name="sizeMode">Способ задания размера.</param>
        /// <param name="widthCells">Ширина в ячейках (для RelativeToCells).</param>
        /// <param name="heightCells">Высота в ячейках (для RelativeToCells).</param>
        void InsertImage(string sheetName, string cell, object source,
            double? widthPx = null, double? heightPx = null,
            ImageSizeMode sizeMode = ImageSizeMode.AbsolutePoints,
            int widthCells = 1, int heightCells = 1);

        // —— Печать и колонтитулы (Этап 5, п.4.6) ——

        /// <summary>Установить область печати: auto — по фактическим данным (UsedRange).</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="auto">Автоподбор (true) или явный адрес (false + printArea).</param>
        /// <param name="printArea">Явный адрес области печати для auto=false.</param>
        void SetPrintArea(string sheetName, bool auto = true, string? printArea = null);

        /// <summary>Прочитать/записать колонтитулы листа (6 слотов).</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="header">Значения верхнего колонтитула (null — не менять слот).</param>
        /// <param name="footer">Значения нижнего колонтитула.</param>
        void SetHeaderFooter(string sheetName, HeaderFooter? header = null, HeaderFooter? footer = null);

        /// <summary>Вернуть колонтитулы листа.</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <returns>Пара header/footer.</returns>
        (HeaderFooter Header, HeaderFooter Footer) GetHeaderFooter(string sheetName);

        /// <summary>Повторяющиеся строки/столбцы при печати («1:2», «A:B»).</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="rows">Диапазон строк или null.</param>
        /// <param name="columns">Диапазон столбцов или null.</param>
        void SetPrintTitles(string sheetName, string? rows, string? columns);

        // —— Листы (Этап 5, п.4.8) ——

        /// <summary>Добавить лист после указанного; при modelSheet — перенести стиль/настройки печати образца.</summary>
        /// <param name="name">Имя нового листа.</param>
        /// <param name="after">Имя листа, после которого вставлять (null — в конец).</param>
        /// <param name="modelSheet">Лист-образец настроек печати (null — без переноса).</param>
        void AddSheet(string name, string? after = null, string? modelSheet = null);

        /// <summary>Удалить лист (нельзя удалить единственный/защищённый).</summary>
        /// <param name="name">Имя листа.</param>
        void RemoveSheet(string name);

        // —— Пагинация (Этапы 8–9, п.5) ——

        /// <summary>Выполнить пагинацию по опциям из <c>WithPagination</c>.</summary>
        /// <returns>Число страниц.</returns>
        int Paginate();

        /// <summary>Выполнить пагинацию с явными опциями (переопределяют конфигурацию билдера).</summary>
        /// <param name="options">Параметры пагинации.</param>
        /// <returns>Число страниц.</returns>
        int Paginate(PaginationOptions options);

        // —— PDF (Этап 7, п.4.7) ——

        /// <summary>Экспортировать один лист в PDF.</summary>
        /// <param name="sheetName">Имя листа.</param>
        /// <param name="pdfPath">Целевой путь .pdf.</param>
        /// <param name="overwrite">Перезаписывать существующий файл.</param>
        /// <param name="quality">Качество экспорта.</param>
        /// <param name="includeMarkup">Включать разметку (IgnorePrintAreas=false и т.п.).</param>
        void ExportSheetToPdf(string sheetName, string pdfPath, bool overwrite = true,
            PdfQuality quality = PdfQuality.Standard, bool includeMarkup = false);

        // —— Сохранение (Этап 2, п.4.4) ——

        /// <summary>Сохранить в целевой путь из WithHiddenWorkFile (SaveAs, формат по расширению).</summary>
        void Save();

        /// <summary>Сохранить в указанный путь (формат по расширению .xls/.xlsx/.xlsm).</summary>
        /// <param name="path">Целевой путь.</param>
        void SaveTo(string path);

        /// <summary>Рабочий файл ещё не сохранён в целевой путь (проверка для тестов/DI).</summary>
        bool IsSaved { get; }
    }

    /// <summary>Шесть слотов колонтитула PageSetup (Left/Center/Right одного уровня).</summary>
    public sealed class HeaderFooter
    {
        /// <summary>Левая часть.</summary>
        public string? Left { get; set; }

        /// <summary>Центральная часть.</summary>
        public string? Center { get; set; }

        /// <summary>Правая часть.</summary>
        public string? Right { get; set; }
    }
}
