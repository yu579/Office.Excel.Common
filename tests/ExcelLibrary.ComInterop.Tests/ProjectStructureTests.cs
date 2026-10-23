using Xunit;

namespace ExcelLibrary.ComInterop.Tests
{
    /// <summary>Смоук-тесты структуры решения (Этап 0). Тесты с реальным Excel будут добавлены позже.</summary>
    public class ProjectStructureTests
    {
        [Fact]
        public void ReferencedLibrary_IsPresent_And_HasExpectedName()
        {
            var lib = typeof(AssemblyInfoMarker).Assembly;
            Assert.Equal("ExcelLibrary.ComInterop", lib.GetName().Name);
        }

        [Fact]
        public void Library_Version_IsDefined()
        {
            Assert.Equal("0.1.0", AssemblyInfoMarker.Version);
        }
    }
}
