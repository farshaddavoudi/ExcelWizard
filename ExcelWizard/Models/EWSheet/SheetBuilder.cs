using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWTable;
using System;
using System.Collections.Generic;
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

    public IExpectSetComponentsSheetBuilder SetTables(ITableBuilder tableBuilder, params ITableBuilder[] tableBuilders)
    {
        ITableBuilder[] tables = new[] { tableBuilder }.Concat(tableBuilders).ToArray();

        Sheet.SheetTables.AddRange(tables.Select(t => (Table)t));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetTables(List<ITableBuilder> tableBuilders)
    {
        Sheet.SheetTables.AddRange(tableBuilders.Select(t => (Table)t));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRows(IRowBuilder rowBuilder, params IRowBuilder[] rowBuilders)
    {
        IRowBuilder[] rows = new[] { rowBuilder }.Concat(rowBuilders).ToArray();

        Sheet.SheetRows.AddRange(rows.Select(r => (Row)r));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetRows(List<IRowBuilder> rowBuilders)
    {
        Sheet.SheetRows.AddRange(rowBuilders.Select(r => (Row)r));

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCells(ICellBuilder cellBuilder, params ICellBuilder[] cellBuilders)
    {
        ICellBuilder[] cells = new[] { cellBuilder }.Concat(cellBuilders).ToArray();

        Sheet.SheetCells.AddRange(cells.Select(c => (Cell)c).ToList());

        return this;
    }

    public IExpectSetComponentsSheetBuilder SetCells(List<ICellBuilder> cellBuilders)
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