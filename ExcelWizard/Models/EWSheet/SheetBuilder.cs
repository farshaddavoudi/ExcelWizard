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

    public IExpectSetComponentsSheetBuilder SetTable(ITableBuilder tableBuilder)
    {
        Sheet.SheetTables.Add((Table)tableBuilder);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetTables(params ITableBuilder[] tables)
    {
        Sheet.SheetTables.AddRange(tables.Select(t => (Table)t));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRow(IRowBuilder rowBuilder)
    {
        Sheet.SheetRows.Add((Row)rowBuilder);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRows(params IRowBuilder[] rowBuilders)
    {
        Sheet.SheetRows.AddRange(rowBuilders.Select(r => (Row)r));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCell(ICellBuilder cellBuilder)
    {
        Sheet.SheetCells.Add((Cell)cellBuilder);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCells(params ICellBuilder[] cellBuilders)
    {
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

    public IExpectOtherPropsAndBuildSheetBuilder SetStyle(SheetStyle sheetStyle)
    {
        Sheet.SheetStyle = sheetStyle;

        CanBuild = true;

        return this;
    }

    public IExpectOtherPropsAndBuildSheetBuilder NoCustomStyle()
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