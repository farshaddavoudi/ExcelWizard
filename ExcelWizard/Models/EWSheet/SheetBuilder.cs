using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWTable;
using System;
using System.Linq;

namespace ExcelWizard.Models.EWSheet;

public class SheetBuilder : IExpectSetComponentsSheetBuilder,
    IExpectStyleSheetBuilder, IExpectOtherPropsAndBuildSheetBuilder
{
    private SheetBuilder() { }

    private Sheet Sheet { get; set; } = new();
    private bool CanBuild { get; set; }

    /// <summary>
    /// Set the Sheet name
    /// </summary>
    /// <param name="sheetName"> Name of the Sheet </param>
    public static IExpectSetComponentsSheetBuilder SetName(string? sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet name cannot be empty");

        return new SheetBuilder
        {
            Sheet = new Sheet
            {
                SheetName = sheetName
            }
        };
    }

    public IExpectSetComponentsSheetBuilder SetTables(params ITableBuilder[] tableBuilders)
    {
        if (tableBuilders.Length == 0)
            throw new ArgumentException("At-least one TableBuilder should be provided for SheetBuilder's SetTables method argument");

        Sheet.SheetTables.AddRange(tableBuilders.Select(t => (Table)t));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRows(params IRowBuilder[] rowBuilders)
    {
        if (rowBuilders.Length == 0)
            throw new ArgumentException("At-least one RowBuilder should be provided for SheetBuilder's SetRows method argument");

        Sheet.SheetRows.AddRange(rowBuilders.Select(r => (Row)r));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCells(params ICellBuilder[] cellBuilders)
    {
        if (cellBuilders.Length == 0)
            throw new ArgumentException("At-least one CellBuilder should be provided for SheetBuilder's SetCells method argument");

        Sheet.SheetCells.AddRange(cellBuilders.Select(c => (Cell)c).ToList());

        return this;
    }

    public IExpectStyleSheetBuilder NoMoreTablesRowsOrCells()
    {
        return this;
    }

    public IExpectOtherPropsAndBuildSheetBuilder SetMergedCells(params IMergeBuilder[] mergeBuilders)
    {
        Sheet.MergedCellsList = mergeBuilders.Select(m => (MergedCells)m).ToList();

        return this;
    }

    public IExpectOtherPropsAndBuildSheetBuilder SetSheetStyle(SheetStyle sheetStyle)
    {
        Sheet.SheetStyle = sheetStyle;

        CanBuild = true;

        return this;
    }

    public IExpectOtherPropsAndBuildSheetBuilder SheetHasNoCustomStyle()
    {
        CanBuild = true;

        return this;
    }

    public IExpectOtherPropsAndBuildSheetBuilder SetLockedStatus(bool isLocked)
    {
        Sheet.IsSheetLocked = isLocked;

        return this;
    }

    public IExpectOtherPropsAndBuildSheetBuilder SetProtectionLevel(ProtectionLevel protectionLevel)
    {
        Sheet.SheetProtectionLevel = protectionLevel;

        return this;
    }

    public ISheetBuilder Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Sheet because some necessary information are not provided");

        return Sheet;
    }
}