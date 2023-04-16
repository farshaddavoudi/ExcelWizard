using Bogus;
using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Tests.Models.EWExcel;
using ValidationException = System.ComponentModel.DataAnnotations.ValidationException;

namespace ExcelWizard.Tests.Models.EWGridLayout;

public class GridExcelSheetTests
{
    [Fact]
    public void ValidateGridExcelSheetInstance_GivenNotIEnumerableDataList_ShouldThrowValidationException()
    {
        // Arrange
        var gridExcelSheet = new GridExcelSheet
        {
            SheetName = new Faker().System.FileName(),
            DataList = new Student()
        };

        // Act
        var action = () => gridExcelSheet.ValidateGridExcelSheetInstance();

        // Assert
        action.Should().Throw<ValidationException>().WithMessage("Object provided for Sheet binding should be a collection of records*");
    }
}