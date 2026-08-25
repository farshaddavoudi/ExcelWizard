using ClosedXML.Excel;
using ExcelWizard.Models;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWSheet;

namespace ExcelWizard.IntegrationTests;

public class CellContentTypeWorkbookTests
{
    [Fact]
    public void AllSupportedCellContentTypes_RoundTripWithExpectedExcelTypes()
    {
        var date = new DateTime(2026, 8, 25, 10, 30, 0, DateTimeKind.Unspecified);
        var cells = new[]
        {
            CreateCell(1, "general", CellContentType.General),
            CreateCell(2, 42.5m, CellContentType.Number),
            CreateCell(3, 1234.6m, CellContentType.Currency),
            CreateCell(4, date, CellContentType.GregorianDateTime),
            CreateCell(5, 123, CellContentType.Text),
            CreateCell(6, 25, CellContentType.Percentage),
            CreateCell(7, "B1*2", CellContentType.Formula),
            CreateCell(8, true, CellContentType.General),
            CreateCell(9, null, CellContentType.General)
        };

        using var workbook = WorkbookTestHelper.GenerateWorkbook(BuildWorkbook(cells));
        var worksheet = workbook.Worksheet("Cell Types");

        worksheet.Cell("A1").DataType.Should().Be(XLDataType.Text);
        worksheet.Cell("A1").GetString().Should().Be("general");
        worksheet.Cell("B1").DataType.Should().Be(XLDataType.Number);
        worksheet.Cell("B1").GetDouble().Should().Be(42.5);
        worksheet.Cell("C1").DataType.Should().Be(XLDataType.Number);
        worksheet.Cell("C1").GetDouble().Should().Be(1235);
        worksheet.Cell("C1").Style.NumberFormat.Format.Should().Be("#,##0");
        worksheet.Cell("D1").DataType.Should().Be(XLDataType.DateTime);
        worksheet.Cell("D1").GetDateTime().Should().Be(date);
        worksheet.Cell("E1").DataType.Should().Be(XLDataType.Text);
        worksheet.Cell("E1").GetString().Should().Be("123");
        worksheet.Cell("F1").DataType.Should().Be(XLDataType.Text);
        worksheet.Cell("F1").GetString().Should().Be("25%");
        worksheet.Cell("G1").HasFormula.Should().BeTrue();
        worksheet.Cell("G1").FormulaA1.Should().Be("B1*2");
        worksheet.Cell("H1").DataType.Should().Be(XLDataType.Boolean);
        worksheet.Cell("H1").GetBoolean().Should().BeTrue();
        worksheet.Cell("I1").IsEmpty().Should().BeTrue();
    }

    [Fact]
    public void NumberContentType_ParsesNumericText()
    {
        var cell = CreateCell(1, "123.5", CellContentType.Number);

        using var workbook = WorkbookTestHelper.GenerateWorkbook(BuildWorkbook(new[] { cell }));

        workbook.Worksheet("Cell Types").Cell("A1").GetDouble().Should().Be(123.5);
    }

    [Fact]
    public void NumberContentType_RejectsNonNumericValues()
    {
        var cell = CreateCell(1, "not-a-number", CellContentType.Number);

        var action = () => WorkbookTestHelper.GenerateWorkbook(BuildWorkbook(new[] { cell }));

        action.Should().Throw<Exception>().WithMessage("*Number CellType*");
    }

    [Fact]
    public void CurrencyContentType_RejectsNonNumericValues()
    {
        var cell = CreateCell(1, "not-currency", CellContentType.Currency);

        var action = () => WorkbookTestHelper.GenerateWorkbook(BuildWorkbook(new[] { cell }));

        action.Should().Throw<Exception>().WithMessage("*Currency CellType*");
    }

    [Fact]
    public void GregorianDateTimeContentType_RejectsNonDateValues()
    {
        var cell = CreateCell(1, "2026-08-25", CellContentType.GregorianDateTime);

        var action = () => WorkbookTestHelper.GenerateWorkbook(BuildWorkbook(new[] { cell }));

        action.Should().Throw<Exception>().WithMessage("*GregorianDateTime CellType*");
    }

    private static Cell CreateCell(int column, object? value, CellContentType contentType)
    {
        return CellBuilder.SetLocation(column, 1)
            .SetValue(value)
            .SetContentType(contentType)
            .Build();
    }

    private static IExcelBuilder BuildWorkbook(IEnumerable<Cell> cells)
    {
        var row = RowBuilder
            .SetCells(cells.Cast<ICellBuilder>().ToList())
            .RowHasNoMerging()
            .RowHasNoCustomStyle()
            .Build();
        var sheet = SheetBuilder
            .SetName("Cell Types")
            .SetRows(row)
            .NoMoreTablesRowsOrCells()
            .SheetHasNoCustomStyle()
            .Build();

        return ExcelBuilder
            .SetGeneratedFileName("cell-types")
            .CreateComplexLayoutExcel()
            .SetSheets(sheet)
            .SheetsHaveNoDefaultStyle()
            .Build();
    }
}
