using Bogus;
using ExcelWizard.Models;
using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Models.EWSheet;
using ValidationException = System.ComponentModel.DataAnnotations.ValidationException;

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

        var givenFileName = new Faker().System.FileName();

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(givenFileName)
            .CreateComplexLayoutExcel()
            .SetSheets(sheetBuilder)
            .SheetsHaveNoDefaultStyle()
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Should().NotBeNull();

        excelModel.GeneratedFileName.Should().Be(givenFileName);
    }

    [Fact]
    public void AddBoundSheet_WhenGivenABindingListModel_ShouldReturnCorrectExcelModelWithOneSheet()
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
            .WithDataBinding()
            .AddBoundSheet(students)
            .Build();

        var excelModel = (ExcelModel)excelBuilder;

        // Assert
        excelModel.Should().NotBeNull();

        excelModel.Sheets.Should().HaveCount(1);

        excelModel.Sheets.First().SheetTables.Should().HaveCount(1);
    }

    [Fact]
    public void AddAnotherBoundSheet_WhenGivenAnotherBoundSheet_ShouldReturnCorrectExcelModelWithTwoSheets()
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

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(new Faker().System.FileName())
            .CreateGridLayoutExcel()
            .WithDataBinding()
            .AddBoundSheet(classAStudents, firstSheetName)
            .AddAnotherBoundSheet(classBStudents, secondSheetName)
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
    public void AddBoundSheets_WhenGivenTwoBoundSheets_ShouldReturnCorrectExcelModelWithTwoSheets()
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

        var boundSheets = new List<BoundSheet>
        {
            new(classAStudents, firstSheetName),
            new(classBStudents, secondSheetName)
        };

        // Act
        var excelBuilder = ExcelBuilder.SetGeneratedFileName(new Faker().System.FileName())
            .CreateGridLayoutExcel()
            .WithDataBinding()
            .AddBoundSheets(boundSheets)
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
    public void AddBoundSheets_WhenGivenASheetWithNotIEnumerableBoundData_ShouldThrow()
    {
        // Arrange
        var studentFaker = new Faker<Student>()
            .RuleFor(x => x.Id, x => x.Random.Int(1, 20))
            .RuleFor(x => x.FullName, x => x.Name.FullName())
            .RuleFor(x => x.Nationality, x => x.Address.Country())
            .RuleFor(x => x.StudentCode, x => x.Random.String(10, 15));

        var classAStudents = studentFaker.Generate(5);

        var classBStudents = new Student();

        var firstSheetName = new Faker().Name.JobArea();
        var secondSheetName = new Faker().Name.JobArea();

        var boundSheets = new List<BoundSheet>
        {
            new(classAStudents, firstSheetName),
            new(classBStudents, secondSheetName)
        };

        // Act
        var act = () => ExcelBuilder.SetGeneratedFileName(new Faker().System.FileName())
            .CreateGridLayoutExcel()
            .WithDataBinding()
            .AddBoundSheets(boundSheets)
            .SheetsHaveNoDefaultStyle()
            .Build();


        // Assert
        act.Should().Throw<ValidationException>().WithMessage("Object provided for Sheet binding should be a collection of records*");
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