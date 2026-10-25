using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelLibrary.ComInterop
{
    /// <summary>
    /// Утилиты нотации A1: разбор/формирование адресов ячеек и диапазонов,
    /// поддержка абсолютных ссылок с «$» (Этапы 3, 8, 9).
    /// </summary>
    public static class A1Address
    {
        private static readonly Regex CellPattern =
            new Regex(@"^\$?(?<col>[A-Za-z]{1,3})\$?(?<row>\d{1,7})$", RegexOptions.Compiled);

        private static readonly Regex SourceRefPattern =
            new Regex(@"(?<![\w!$:])\$?(?:'[^']+'|[A-Za-z_А-Яа-яЁё\.]+)!\$?(?<col>[A-Za-z]{1,3})\$?(?<row>\d{1,7})",
                RegexOptions.Compiled);

        /// <summary>Разбирает адрес ячейки («A1», «$AB$123») в пару (столбец, строка), 1-based.</summary>
        /// <param name="address">Адрес ячейки в нотации A1 (допускаются $).</param>
        /// <param name="column">Номер столбца (A=1).</param>
        /// <param name="row">Номер строки.</param>
        /// <returns>true, если адрес корректен.</returns>
        public static bool TryParse(string address, out int column, out int row)
        {
            column = 0;
            row = 0;
            if (string.IsNullOrWhiteSpace(address)) return false;

            var m = CellPattern.Match(address.Trim());
            if (!m.Success) return false;

            column = ColumnNameToIndex(m.Groups["col"].Value);
            row = int.Parse(m.Groups["row"].Value, CultureInfo.InvariantCulture);
            return column >= 1 && row >= 1;
        }

        /// <summary>Разбирает адрес ячейки; при некорректном адресе — <see cref="FormatException"/>.</summary>
        /// <param name="address">Адрес ячейки.</param>
        /// <returns>(столбец, строка), 1-based.</returns>
        public static (int Column, int Row) Parse(string address)
        {
            if (!TryParse(address, out int c, out int r))
                throw new FormatException($"Некорректный адрес ячейки A1: '{address}'.");
            return (c, r);
        }

        /// <summary>Формирует адрес ячейки из номеров столбца и строки (1-based).</summary>
        /// <param name="column">Номер столбца (A=1).</param>
        /// <param name="row">Номер строки.</param>
        /// <returns>Адрес вида «AB123».</returns>
        public static string Format(int column, int row)
        {
            if (column < 1 || row < 1)
                throw new ArgumentOutOfRangeException(nameof(column), "Столбец и строка должны быть >= 1.");
            return IndexToColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Имя столбца по номеру (1 → A, 27 → AA).</summary>
        /// <param name="index">Номер столбца, 1-based.</param>
        /// <returns>Имя столбца.</returns>
        public static string IndexToColumnName(int index)
        {
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));
            string name = string.Empty;
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                name = (char)('A' + rem) + name;
                index = (index - 1) / 26;
            }
            return name;
        }

        /// <summary>Номер столбца по имени (A → 1, AA → 27). Регистр не учитывается.</summary>
        /// <param name="name">Имя столбца (без $).</param>
        /// <returns>Номер столбца, 1-based; 0 при некорректном имени.</returns>
        public static int ColumnNameToIndex(string name)
        {
            if (string.IsNullOrEmpty(name)) return 0;
            int index = 0;
            foreach (char ch in name.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                index = index * 26 + (ch - 'A' + 1);
            }
            return index;
        }

        /// <summary>
        /// Извлекает все ссылки вида <c>=Лист!A11</c> / <c>Лист!$B$12</c> из формулы
        /// (для сканирования шаблона, Этап 8). Возвращает уникальные пары (столбец, строка).
        /// </summary>
        /// <param name="formula">Текст формулы (с «=» или без).</param>
        /// <returns>Перечисление найденных ссылок (столбец, строка).</returns>
        public static IEnumerable<(int Column, int Row)> ExtractSheetRefs(string formula)
        {
            if (string.IsNullOrEmpty(formula)) yield break;
            foreach (Match m in SourceRefPattern.Matches(formula))
            {
                int col = ColumnNameToIndex(m.Groups["col"].Value);
                int row = int.Parse(m.Groups["row"].Value, CultureInfo.InvariantCulture);
                if (col >= 1 && row >= 1)
                    yield return (col, row);
            }
        }

        /// <summary>Проверяет, содержит ли формула ссылку на указанный лист (по имени, без учёта регистра).</summary>
        /// <param name="formula">Текст формулы.</param>
        /// <param name="sheetName">Имя листа-источника.</param>
        /// <returns>true — есть ссылка вида SheetName!A1.</returns>
        public static bool ReferencesSheet(string formula, string sheetName)
        {
            if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(sheetName)) return false;
            // имена листов могут быть в кавычках: 'Шаблон листа'!A1
            string pattern = @"(?<![\w!$:])\$?(?:'" + Regex.Escape(sheetName) + "'|" + Regex.Escape(sheetName) + @")!\$?[A-Za-z]{1,3}\$?\d{1,7}";
            return Regex.IsMatch(formula, pattern, RegexOptions.IgnoreCase);
        }
    }
}
