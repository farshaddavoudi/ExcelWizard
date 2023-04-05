using ExcelWizard.Models;
using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWSheet;

namespace ExcelWizard.Tests.Models.EWExcel;

public class ExcelBuilderTests
{
    [Fact]
    public void SetGeneratedFileName_WhenGivenAFileName_ShouldSetExcelFileName()
    {
        // Arrange
        var sheetBuilder = Mock.Of<Sheet>();

        // Act
        var excel = () => ExcelBuilder.SetGeneratedFileName(null)
            .CreateComplexLayoutExcel()
            .SetSheets(sheetBuilder)
            .SheetsHaveNoDefaultStyle()
            .Build();

        // Assert
        excel.Should().Throw<ArgumentNullException>().WithMessage("Excel file name cannot be empty*");
    }

    [Fact]
    public void SetGeneratedFileName_WhenGivenNoOrEmptyFileName_ShouldThrow()
    {
        // Arrange
        var sheetBuilder = Mock.Of<Sheet>();

        var givenName = "New-Generated-Excel";

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(givenName)
            .CreateComplexLayoutExcel()
            .SetSheets(sheetBuilder)
            .SheetsHaveNoDefaultStyle()
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Should().NotBeNull();

        excelModel.GeneratedFileName.Should().Be(givenName);
    }

    [Fact]
    public void WithOneSheetUsingModelBinding_WhenGivenABindingListModel_ShouldReturnCorrectExcelModelWithOneSheet()
    {
        // Arrange
        var students = new List<Student>
        {
            new()
            {
                Id = 1,
                FullName = "Farshad Davoudi",
                Nationality = "Germany",
                StudentCode = "St32"
            },
            new()
            {
                Id = 2,
                FullName = "Somaye Ebrahimi",
                Nationality = "Iran",
                StudentCode = "St34"
            },
            new()
            {
                Id = 3,
                FullName = "Leonardo Decaprio",
                Nationality = "US",
                StudentCode = "St36"
            }
        };

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName("excel-name")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingModelBinding(students)
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Should().NotBeNull();

        excelModel.Sheets.Should().HaveCount(1);

        excelModel.Sheets.First().SheetTables.Should().HaveCount(1);
    }

    [Fact]
    public void WithMultipleSheetsUsingModelBinding_WhenGivenListOfTwoBindingListModel_ShouldReturnCorrectExcelModelWithTwoSheets()
    {
        // Arrange
        var classAStudents = new List<Student>
        {
            new()
            {
                Id = 1,
                FullName = "Farshad Davoudi",
                Nationality = "Germany",
                StudentCode = "St32"
            },
            new()
            {
                Id = 2,
                FullName = "Somaye Ebrahimi",
                Nationality = "Iran",
                StudentCode = "St34"
            },
            new()
            {
                Id = 3,
                FullName = "Leonardo Decaprio",
                Nationality = "US",
                StudentCode = "St36"
            }
        };

        var classBStudents = new List<Student>
        {
            new()
            {
                Id = 11,
                FullName = "Farshad Davoudi2",
                Nationality = "Germany2",
                StudentCode = "St322"
            },
            new()
            {
                Id = 22,
                FullName = "Somaye Ebrahimi2",
                Nationality = "Iran2",
                StudentCode = "St342"
            },
            new()
            {
                Id = 33,
                FullName = "Leonardo Decaprio2",
                Nationality = "US2",
                StudentCode = "St362"
            }
        };

        var firstSheetName = "Class A";
        var secondSheetName = "Class B";

        var sheets = new List<BindingSheet>
        {
            new(classAStudents, firstSheetName),
            new(classBStudents, secondSheetName)
        };

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName("excel-name")
            .CreateGridLayoutExcel()
            .WithMultipleSheetsUsingModelBinding(sheets)
            .SheetsHaveNoDefaultStyle()
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Should().NotBeNull();

        excelModel.Sheets.Should().HaveCount(2);

        excelModel.Sheets.First().SheetName.Should().Be(firstSheetName);
        excelModel.Sheets.Last().SheetName.Should().Be(secondSheetName);
    }

    [Fact]
    public void SetSheets_WhenGivenOneSheet_ShouldSetSheetCorrectlyOnExcelModel()
    {
        // Arrange
        var sheet = Mock.Of<Sheet>();

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName("excel-name")
            .CreateComplexLayoutExcel()
            .SetSheets(sheet)
            .SheetsHaveNoDefaultStyle()
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Sheets.Should().BeEquivalentTo(new List<Sheet> { sheet });
    }

    [Fact]
    public void SetSheets_WhenGivenThreeSheets_ShouldSetSheetsCorrectlyOnExcelModel()
    {
        // Arrange
        var sheet1 = Mock.Of<Sheet>();
        var sheet2 = Mock.Of<Sheet>();
        var sheet3 = Mock.Of<Sheet>();

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName("excel-name")
            .CreateComplexLayoutExcel()
            .SetSheets(sheet1, sheet2, sheet3)
            .SheetsHaveNoDefaultStyle()
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Sheets.Should().BeEquivalentTo(new List<Sheet> { sheet1, sheet2, sheet3 });
    }
}