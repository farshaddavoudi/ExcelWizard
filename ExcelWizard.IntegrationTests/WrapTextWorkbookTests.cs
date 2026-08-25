using ClosedXML.Excel;
using ExcelWizard.Models;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWSheet;
using ExcelWizard.Models.EWTable;

namespace ExcelWizard.IntegrationTests;

public class WrapTextWorkbookTests
{
    private const string LongComment =
        "این یک توضیح طولانی برای بررسی نمایش کامل متن در چند خط است و محتوای اصلی نباید تغییر کند " +
        "This paragraph also contains enough English words to require several wrapped lines in Excel.";

    [Fact]
    public void GridLayout_WrapsLongTextAndAutomaticallyExpandsTheRow()
    {
        var builder = ExcelBuilder
            .SetGeneratedFileName("grid-wrap")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingModelBinding(new[]
            {
                new WrappedGridRow { Comment = LongComment, SecondaryComment = "line one\nline two\nline three", Code = "A-1" }
            })
            .Build();

        using var workbook = WorkbookTestHelper.GenerateWorkbook(builder);
        var worksheet = workbook.Worksheet("Wrapped Grid");

        worksheet.Cell(1, 1).Style.Alignment.WrapText.Should().BeFalse();
        worksheet.Cell(2, 1).Style.Alignment.WrapText.Should().BeTrue();
        worksheet.Cell(2, 2).Style.Alignment.WrapText.Should().BeTrue();
        worksheet.Cell(2, 3).Style.Alignment.WrapText.Should().BeFalse();
        worksheet.Cell(2, 1).GetString().Should().Be(LongComment);
        worksheet.Cell(2, 2).GetString().Should().Be("line one\nline two\nline three");
        worksheet.Row(2).Height.Should().BeGreaterThan(worksheet.RowHeight);
    }

    [Fact]
    public void GridLayout_PreservesExplicitDataRowHeightWhenWrapping()
    {
        var builder = ExcelBuilder
            .SetGeneratedFileName("grid-explicit-height")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingModelBinding(new[] { new FixedHeightGridRow { Comment = LongComment } })
            .Build();

        using var workbook = WorkbookTestHelper.GenerateWorkbook(builder);
        var worksheet = workbook.Worksheet("Fixed Height");

        worksheet.Cell(2, 1).Style.Alignment.WrapText.Should().BeTrue();
        worksheet.Row(2).Height.Should().Be(42);
    }

    [Fact]
    public void ModelBoundTable_WrapsOnlyConfiguredDataColumns()
    {
        var table = TableBuilder
            .CreateUsingAModelToBind(
                new[] { new WrappedTableRow { Comment = LongComment, Code = "T-1" } },
                new CellLocation(1, 1))
            .BoundTableHasNoMerging()
            .Build();

        var sheet = SheetBuilder
            .SetName("Wrapped Table")
            .SetTables(table)
            .NoMoreTablesRowsOrCells()
            .SheetHasNoCustomStyle()
            .Build();

        var builder = ExcelBuilder
            .SetGeneratedFileName("table-wrap")
            .CreateComplexLayoutExcel()
            .SetSheets(sheet)
            .SheetsHaveNoDefaultStyle()
            .Build();

        using var workbook = WorkbookTestHelper.GenerateWorkbook(builder);
        var worksheet = workbook.Worksheet("Wrapped Table");

        worksheet.Cell(1, 1).Style.Alignment.WrapText.Should().BeFalse();
        worksheet.Cell(2, 1).Style.Alignment.WrapText.Should().BeTrue();
        worksheet.Cell(2, 2).Style.Alignment.WrapText.Should().BeFalse();
        worksheet.Cell(2, 1).GetString().Should().Be(LongComment);
        worksheet.Row(2).Height.Should().BeGreaterThan(worksheet.RowHeight);
    }

    [Fact]
    public void ManualRow_WrapTextAutomaticallyExpandsTheRow()
    {
        var builder = BuildManualRowWorkbook(rowHeight: null);

        using var workbook = WorkbookTestHelper.GenerateWorkbook(builder);
        var worksheet = workbook.Worksheet("Manual Wrap");

        worksheet.Cell("A1").Style.Alignment.WrapText.Should().BeTrue();
        worksheet.Cell("A1").GetString().Should().Be(LongComment);
        worksheet.Row(1).Height.Should().BeGreaterThan(worksheet.RowHeight);
    }

    [Fact]
    public void ManualRow_WordwrapAliasStillEnablesWrapping()
    {
        var cell = CellBuilder.SetLocation("A", 1)
            .SetValue("line one\nline two")
            .SetCellStyle(new CellStyle { Wordwrap = true })
            .Build();
        var builder = BuildManualRowWorkbook(cell, rowHeight: null);

        using var workbook = WorkbookTestHelper.GenerateWorkbook(builder);

        workbook.Worksheet("Manual Wrap").Cell("A1").Style.Alignment.WrapText.Should().BeTrue();
    }

    [Fact]
    public void ManualRow_PreservesExplicitHeightWhenWrapping()
    {
        var builder = BuildManualRowWorkbook(rowHeight: 36);

        using var workbook = WorkbookTestHelper.GenerateWorkbook(builder);
        var worksheet = workbook.Worksheet("Manual Wrap");

        worksheet.Cell("A1").Style.Alignment.WrapText.Should().BeTrue();
        worksheet.Row(1).Height.Should().Be(36);
    }

    [Fact]
    public void UnconfiguredGridColumn_RemainsUnwrappedAtDefaultHeight()
    {
        var builder = ExcelBuilder
            .SetGeneratedFileName("grid-default")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingModelBinding(new[] { new DefaultGridRow { Comment = LongComment } })
            .Build();

        using var workbook = WorkbookTestHelper.GenerateWorkbook(builder);
        var worksheet = workbook.Worksheet("Default Grid");

        worksheet.Cell(2, 1).Style.Alignment.WrapText.Should().BeFalse();
        worksheet.Row(2).Height.Should().Be(worksheet.RowHeight);
    }

    private static IExcelBuilder BuildManualRowWorkbook(double? rowHeight)
    {
        var cell = CellBuilder.SetLocation("A", 1)
            .SetValue(LongComment)
            .SetCellStyle(new CellStyle { WrapText = true })
            .Build();

        return BuildManualRowWorkbook(cell, rowHeight);
    }

    private static IExcelBuilder BuildManualRowWorkbook(Cell cell, double? rowHeight)
    {
        var rowBuilder = RowBuilder
            .SetCells(cell)
            .RowHasNoMerging();
        var row = rowHeight is null
            ? rowBuilder.RowHasNoCustomStyle().Build()
            : rowBuilder.SetRowStyle(new RowStyle { RowHeight = rowHeight }).Build();
        var sheet = SheetBuilder
            .SetName("Manual Wrap")
            .SetRows(row)
            .NoMoreTablesRowsOrCells()
            .SetSheetStyle(new SheetStyle { SheetDefaultColumnWidth = 18 })
            .Build();

        return ExcelBuilder
            .SetGeneratedFileName("manual-wrap")
            .CreateComplexLayoutExcel()
            .SetSheets(sheet)
            .SheetsHaveNoDefaultStyle()
            .Build();
    }

    [ExcelSheet(SheetName = "Wrapped Grid")]
    private sealed class WrappedGridRow
    {
        [ExcelSheetColumn(HeaderName = "Comment", ColumnWidth = 24, WrapText = true)]
        public string? Comment { get; init; }

        [ExcelSheetColumn(HeaderName = "Secondary", ColumnWidth = 18, WrapText = true)]
        public string? SecondaryComment { get; init; }

        [ExcelSheetColumn(HeaderName = "Code", ColumnWidth = 12)]
        public string? Code { get; init; }
    }

    [ExcelSheet(SheetName = "Fixed Height", DataRowHeight = 42)]
    private sealed class FixedHeightGridRow
    {
        [ExcelSheetColumn(ColumnWidth = 18, WrapText = true)]
        public string? Comment { get; init; }
    }

    [ExcelSheet(SheetName = "Default Grid")]
    private sealed class DefaultGridRow
    {
        [ExcelSheetColumn(ColumnWidth = 18)]
        public string? Comment { get; init; }
    }

    [ExcelTable]
    private sealed class WrappedTableRow
    {
        [ExcelTableColumn(HeaderName = "Comment", WrapText = true)]
        public string? Comment { get; init; }

        public string? Code { get; init; }
    }
}
