using System;
using System.Runtime.Serialization;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Базовое исключение библиотеки ExcelLibrary.ComInterop.
    /// Все специфичные ошибки библиотеки наследуются от этого типа,
    /// что позволяет перехватывать любые ошибки одним <c>catch (ExcelLibraryException)</c>.
    /// </summary>
    [Serializable]
    public class ExcelLibraryException : Exception
    {
        /// <summary>Инициализирует новый экземпляр с заданным сообщением.</summary>
        /// <param name="message">Текст ошибки.</param>
        public ExcelLibraryException(string message) : base(message) { }

        /// <summary>Инициализирует новый экземпляр с сообщением и внутренним исключением.</summary>
        /// <param name="message">Текст ошибки.</param>
        /// <param name="inner">Внутреннее (исходное) исключение.</param>
        public ExcelLibraryException(string message, Exception inner) : base(message, inner) { }

        /// <summary>Конструктор десериализации (для поддержки бинарной сериализации исключений).</summary>
        protected ExcelLibraryException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }

    /// <summary>
    /// Ошибка переполнения диапазона: данные не помещаются в доступную область листа
    /// (например, запись за пределы строки 1 048 576 для .xlsx или 65 536 строк для .xls),
    /// либо объём данных превышает ёмкость страниц шаблона при неразрешимой конфигурации.
    /// </summary>
    [Serializable]
    public class ExcelRangeOverflowException : ExcelLibraryException
    {
        /// <summary>Количество элементов, которое пыталась записать библиотека.</summary>
        public int RequestedCount { get; }

        /// <summary>Максимально допустимое количество элементов в данной конфигурации.</summary>
        public int MaximumCount { get; }

        /// <summary>Инициализирует новый экземпляр <see cref="ExcelRangeOverflowException"/>.</summary>
        /// <param name="message">Текст ошибки.</param>
        /// <param name="requestedCount">Запрошенное количество элементов.</param>
        /// <param name="maximumCount">Максимально допустимое количество элементов.</param>
        public ExcelRangeOverflowException(string message, int requestedCount = 0, int maximumCount = 0)
            : base(message)
        {
            RequestedCount = requestedCount;
            MaximumCount = maximumCount;
        }

        /// <summary>Инициализирует новый экземпляр с внутренним исключением.</summary>
        /// <param name="message">Текст ошибки.</param>
        /// <param name="inner">Внутреннее исключение.</param>
        /// <param name="requestedCount">Запрошенное количество элементов.</param>
        /// <param name="maximumCount">Максимально допустимое количество элементов.</param>
        public ExcelRangeOverflowException(string message, Exception inner, int requestedCount = 0, int maximumCount = 0)
            : base(message, inner)
        {
            RequestedCount = requestedCount;
            MaximumCount = maximumCount;
        }

    }

    /// <summary>
    /// Ошибка шаблона Excel: отсутствует обязательный лист, именованный диапазон,
    /// маркер пагинации или структура шаблона не соответствует ожиданиям библиотеки.
    /// </summary>
    [Serializable]
    public class ExcelTemplateException : ExcelLibraryException
    {
        /// <summary>Имя проблемного элемента шаблона (лист, диапазон, ячейка).</summary>
        public string? ElementName { get; }

        /// <summary>Инициализирует новый экземпляр <see cref="ExcelTemplateException"/>.</summary>
        /// <param name="message">Текст ошибки.</param>
        /// <param name="elementName">Имя проблемного элемента шаблона.</param>
        public ExcelTemplateException(string message, string? elementName = null) : base(message)
        {
            ElementName = elementName;
        }

        /// <summary>Инициализирует новый экземпляр с внутренним исключением.</summary>
        /// <param name="message">Текст ошибки.</param>
        /// <param name="inner">Внутреннее исключение.</param>
        /// <param name="elementName">Имя проблемного элемента шаблона.</param>
        public ExcelTemplateException(string message, Exception inner, string? elementName = null) : base(message, inner)
        {
            ElementName = elementName;
        }
    }

    /// <summary>
    /// Несовпадение битности процесса (.NET) и установленного Microsoft Excel.
    /// COM-взаимодействие возможно только между процессами одинаковой разрядности (x86↔x86, x64↔x64).
    /// В сообщении содержится инструкция по устранению.
    /// </summary>
    [Serializable]
    public class ExcelBitnessMismatchException : ExcelLibraryException
    {
        /// <summary>Битность текущего процесса, байт (32 или 64).</summary>
        public int ProcessBitness { get; }

        /// <summary>Битность установленного Excel, байт (32 или 64), если удалось определить.</summary>
        public int? ExcelBitness { get; }

        /// <summary>Инициализирует новый экземпляр <see cref="ExcelBitnessMismatchException"/>.</summary>
        /// <param name="processBitness">Битность процесса.</param>
        /// <param name="excelBitness">Битность Excel (или null, если не определена).</param>
        public ExcelBitnessMismatchException(int processBitness, int? excelBitness)
            : base(BuildMessage(processBitness, excelBitness))
        {
            ProcessBitness = processBitness;
            ExcelBitness = excelBitness;
        }

        private static string BuildMessage(int processBitness, int? excelBitness)
        {
            string excelText = excelBitness.HasValue ? excelBitness.Value + "-бит" : "не определена";
            return
                $"Несовпадение битности: текущий процесс — {processBitness}-бит, установленный Excel — {excelText}. " +
                "COM-взаимодействие с Excel требует одинаковой разрядности процесса и Office. " +
                "Способы устранения: (1) собрать приложение под платформу, совпадающую с Office " +
                "(AnyCPU с включённым 'Prefer 32-bit' для 32-битного Office или x64 для 64-битного); " +
                "(2) использовать NuGet-пакет ExcelLibrary с подборкой сборки x86/x64 через рантайм-селектор; " +
                "(3) установить Office той же разрядности, что и процесс.";
        }
    }

    /// <summary>
    /// Ошибка автоматической пагинации: шаблон не содержит распознаваемых параметров
    /// разбиения (маркеры, ссылки источника), расчёт числа страниц невозможен или
    /// процесс разбиения был прерван (выполнен откат).
    /// </summary>
    [Serializable]
    public class ExcelPaginationException : ExcelLibraryException
    {
        /// <summary>Инициализирует новый экземпляр <see cref="ExcelPaginationException"/>.</summary>
        /// <param name="message">Текст ошибки.</param>
        public ExcelPaginationException(string message) : base(message) { }

        /// <summary>Инициализирует новый экземпляр с внутренним исключением.</summary>
        /// <param name="message">Текст ошибки.</param>
        /// <param name="inner">Внутреннее исключение.</param>
        public ExcelPaginationException(string message, Exception inner) : base(message, inner) { }
    }
}
