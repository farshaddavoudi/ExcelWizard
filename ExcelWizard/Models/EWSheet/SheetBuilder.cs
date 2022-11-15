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

    public IExpectSetComponentsSheetBuilder SetTable(ITableBuilder table)
    {
        Sheet.SheetTables.Add((Table)table);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetTables(params ITableBuilder[] tables)
    {
        Sheet.SheetTables.AddRange(tables.Select(t => (Table)t));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRow(IRowBuilder row)
    {
        Sheet.SheetRows.Add((Row)row);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRows(params IRowBuilder[] rows)
    {
        Sheet.SheetRows.AddRange(rows.Select(r => (Row)r));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCell(ICellBuilder cell)
    {
        Sheet.SheetCells.Add((Cell)cell);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCells(params ICellBuilder[] cells)
    {
        Sheet.SheetCells.AddRange(cells.Select(c => (Cell)c).ToList());

        return this;
    }

    public IExpectStyleSheetBuilder NoMoreTablesRowsOrCells()
    {
        return this;
    }

    public IExpectOtherPropsAndBuildSheetBuilder SetMergedCells(params IMergeBuilder[] merges)
    {
        Sheet.MergedCellsList = merges.Select(m => (MergedCells)m).ToList();

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

    public Sheet Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Sheet because some necessary information are not provided");

        return Sheet;
    }
}