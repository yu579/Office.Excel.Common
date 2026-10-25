using System;

using System.Runtime.CompilerServices;

// Тестовому сборке открываем internal-члены COM-слоя (GetAppProperty и пр.) для проверки сценариев Этапа 1.
[assembly: InternalsVisibleTo("ExcelLibrary.ComInterop.Tests")]
[assembly: CLSCompliant(false)] // COM Interop требует не-CLS-совместимых типов (dynamic/COM)
