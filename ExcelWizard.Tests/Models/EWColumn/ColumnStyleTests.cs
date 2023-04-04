using ExcelWizard.Models.EWColumn;

namespace ExcelWizard.Tests.Models.EWColumn
{
    public class ColumnStyleTests
    {
        [Fact]
        public void NewInstance_ColumnTextAlignShouldbeCenter()
        {
            // Arrange
            var colStyle = new ColumnStyle("A");

            // Assert
            colStyle.ColumnTextAlign.Should().Be(ExcelWizard.Models.EWStyles.TextAlign.Center);
        }

        [Fact]
        public void NewInstance_WhenProvideLetterColumnInConstructor_ShouldSetEquivalentColumn()
        {
            // Arrange
            var colStyle = new ColumnStyle("B");

            // Assert
            colStyle.ColumnNumber.Should().Be(2);
        }
    }
}
