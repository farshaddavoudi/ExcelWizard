using ExcelWizard.Models;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Models.EWTable;

namespace ExcelWizard.Tests.Models;

public class WrapTextTests
{
    [Fact]
    public void CellStyle_DefaultsToNotWrapped()
    {
        var style = new CellStyle();

        style.WrapText.Should().BeFalse();
        style.Wordwrap.Should().BeFalse();
    }

    [Fact]
    public void CellStyle_WordwrapAliasUpdatesWrapText()
    {
        var style = new CellStyle { Wordwrap = true };

        style.WrapText.Should().BeTrue();
    }

    [Fact]
    public void CellStyle_WrapTextUpdatesWordwrapAlias()
    {
        var style = new CellStyle { WrapText = true };

        style.Wordwrap.Should().BeTrue();
    }

    [Fact]
    public void ColumnAttributes_DefaultToNotWrapped()
    {
        new ExcelSheetColumnAttribute().WrapText.Should().BeFalse();
        new ExcelTableColumnAttribute().WrapText.Should().BeFalse();
    }

    [Fact]
    public void GridBuilder_MapsWrapTextToDataCellOnly()
    {
        var excel = (ExcelModel)ExcelBuilder
            .SetGeneratedFileName("grid-wrap")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingModelBinding(new[] { new GridRow { Comment = "long comment", Code = "A1" } })
            .Build();

        var sheet = excel.Sheets.Single();
        var headerRow = sheet.SheetRows.Single();
        var dataRow = sheet.SheetTables.Single().TableRows.Single();

        headerRow.RowCells.Should().OnlyContain(cell => !cell.CellStyle.WrapText);
        dataRow.RowCells.Single(cell => Equals(cell.CellValue, "long comment")).CellStyle.WrapText.Should().BeTrue();
        dataRow.RowCells.Single(cell => Equals(cell.CellValue, "A1")).CellStyle.WrapText.Should().BeFalse();
    }

    [Fact]
    public void TableBuilder_MapsWrapTextToDataCellOnly()
    {
        var table = (Table)TableBuilder
            .CreateUsingAModelToBind(
                new[] { new TableRow { Comment = "long comment", Code = "A1" } },
                new CellLocation(1, 1))
            .BoundTableHasNoMerging()
            .Build();

        var headerRow = table.TableRows.First();
        var dataRow = table.TableRows.Last();

        headerRow.RowCells.Should().OnlyContain(cell => !cell.CellStyle.WrapText);
        dataRow.RowCells.Single(cell => Equals(cell.CellValue, "long comment")).CellStyle.WrapText.Should().BeTrue();
        dataRow.RowCells.Single(cell => Equals(cell.CellValue, "A1")).CellStyle.WrapText.Should().BeFalse();
    }

    [ExcelSheet]
    private sealed class GridRow
    {
        [ExcelSheetColumn(WrapText = true)]
        public string? Comment { get; init; }

        public string? Code { get; init; }
    }

    [ExcelTable]
    private sealed class TableRow
    {
        [ExcelTableColumn(WrapText = true)]
        public string? Comment { get; init; }

        public string? Code { get; init; }
    }
}
