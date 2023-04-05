using ExcelWizard.Models;
using System.ComponentModel.DataAnnotations;

namespace ExcelWizard.Tests.Models.EWColumn
{
    public class ColumnWidthTests
    {
        [Fact]
        public void Validate_WhenSupplyWidthWhileWidthCalcTypeIsSetToAdjustToContent_ShouldReturnErrorInValidationResult()
        {
            // Arrange
            var model = new ColumnWidth { Width = 200 };

            var validationContext = new ValidationContext(model);

            // Act
            var result = model.Validate(validationContext).ToList();

            // Assert
            result.Should().NotBeNull();

            result.Should().HaveCount(1);

            result.First().ErrorMessage.Should().BeEquivalentTo("Column with AdjustToContent Width calculation type cannot have explicit value");
        }

        [Fact]
        public void Validate_WhenNotSupplyWidthWhileWidthCalcTypeIsSetToExplicitValue_ShouldReturnErrorInValidationResult()
        {
            // Arrange
            var model = new ColumnWidth { WidthCalculationType = ColumnWidthCalculationType.ExplicitValue };

            var validationContext = new ValidationContext(model);

            // Act
            var result = model.Validate(validationContext).ToList();

            // Assert
            result.Should().NotBeNull();

            result.Should().HaveCount(1);

            result.First().ErrorMessage.Should().BeEquivalentTo("Column width value should be specified when CalculationType is set to explicit value");
        }
    }
}
