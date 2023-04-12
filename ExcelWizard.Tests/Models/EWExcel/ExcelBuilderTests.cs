using Bogus;
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

        var givenName = new Faker().System.FileName();

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
        var studentFaker = new Faker<Student>()
            .RuleFor(x => x.Id, x => x.Random.Int(1, 20))
            .RuleFor(x => x.FullName, x => x.Name.FullName())
            .RuleFor(x => x.Nationality, x => x.Address.Country())
            .RuleFor(x => x.StudentCode, x => x.Random.String(10, 15));

        var students = studentFaker.Generate(4);

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(new Faker().System.FileName())
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
        var studentFaker = new Faker<Student>()
            .RuleFor(x => x.Id, x => x.Random.Int(1, 20))
            .RuleFor(x => x.FullName, x => x.Name.FullName())
            .RuleFor(x => x.Nationality, x => x.Address.Country())
            .RuleFor(x => x.StudentCode, x => x.Random.String(10, 15));

        var classAStudents = studentFaker.Generate(5);

        var classBStudents = studentFaker.Generate(4);

        var firstSheetName = new Faker().Name.JobArea();
        var secondSheetName = new Faker().Name.JobArea();

        var sheets = new List<BindingSheet>
        {
            new(classAStudents, firstSheetName),
            new(classBStudents, secondSheetName)
        };

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(new Faker().System.FileName())
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
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(new Faker().System.FileName())
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
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(new Faker().System.FileName())
            .CreateComplexLayoutExcel()
            .SetSheets(sheet1, sheet2, sheet3)
            .SheetsHaveNoDefaultStyle()
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Sheets.Should().BeEquivalentTo(new List<Sheet> { sheet1, sheet2, sheet3 });
    }
}