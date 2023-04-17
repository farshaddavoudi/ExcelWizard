using Bogus;
using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Tests.Models.EWExcel;
using ValidationException = System.ComponentModel.DataAnnotations.ValidationException;

namespace ExcelWizard.Tests.Models.EWGridLayout;

public class GridExcelSheetTests
{
    [Fact]
    public void ValidateBoundSheetInstance_GivenNotIEnumerableBoundData_ShouldThrowValidationException()
    {
        // Arrange
        var boundSheet = new BoundSheet(new Student(), new Faker().System.FileName());

        // Act
        var action = () => boundSheet.ValidateBoundSheetInstance();

        // Assert
        action.Should().Throw<ValidationException>().WithMessage("Object provided for Sheet binding should be a collection of records*");
    }
}