using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWTable;
using System;
using System.Collections.Generic;

namespace ExcelWizard.Models.EWSheet;

public class SheetBuilder : ISheetBuilder, IExpectSetComponentsSheetBuilder,
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

    public IExpectSetComponentsSheetBuilder SetTable(Table table)
    {
        Sheet.SheetTables.Add(table);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetTables(List<Table> tables)
    {
        Sheet.SheetTables.AddRange(tables);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRow(Row row)
    {
        Sheet.SheetRows.Add(row);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRows(List<Row> rows)
    {
        Sheet.SheetRows.AddRange(rows);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCell(Cell cell)
    {
        Sheet.SheetCells.Add(cell);

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCells(List<Cell> cells)
    {
        Sheet.SheetCells.AddRange(cells);

        return this;
    }

    public IExpectStyleSheetBuilder NoMoreTablesRowsOrCells()
    {
        return this;
    }

    public IExpectStyleSheetBuilder SetMergedCells(List<string> mergedCells)
    {
        Sheet.MergedCells = mergedCells;

        return this;
    }

    public IExpectStyleSheetBuilder IgnoreMerges()
    {
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

    public SheetBuilder SetLockedStatus(bool isLocked)
    {
        Sheet.IsSheetLocked = isLocked;

        return this;
    }

    public SheetBuilder SetProtectionLevel(ProtectionLevel protectionLevel)
    {
        Sheet.SheetProtectionLevel = protectionLevel;

        return this;
    }

    public Sheet Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Sheet because some necessary information not provided");

        return Sheet;
    }
}