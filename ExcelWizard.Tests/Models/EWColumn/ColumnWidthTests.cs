using ExcelWizard.Models;

namespace ExcelWizard.Tests.Models.EWColumn
{
    public class ColumnWidthTests
    {
        [Fact]
        public void Validate_WhenSupplyWidthWhileWidthCalcTypeIsSetToAdjustToContent_ShouldReturnValidationResult()
        {
            // Arrange
            var a = new ColumnWidth { Width = 200 };

            // Act
            var act = a.Validate;

            // Assert
            act.Should().a
        }
    }
}
