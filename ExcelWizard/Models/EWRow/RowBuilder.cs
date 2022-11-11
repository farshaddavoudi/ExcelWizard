using ExcelWizard.Models.EWCell;
using System;
using System.Collections.Generic;

namespace ExcelWizard.Models.EWRow;

public class RowBuilder : IExpectMergedCellsStatusRowBuilder, IRowBuilder
{
    private RowBuilder() { }

    private Row Row { get; set; } = new();
    private bool CanBuild { get; set; }

    /// <summary>
    /// Each Row contains one or more Cell(s). It is required as Row definition cannot be without Cells.
    /// </summary>
    public static IExpectMergedCellsStatusRowBuilder SetCells(List<Cell> rowCells)
    {
        return new RowBuilder
        {
            Row = new Row
            {
                RowCells = rowCells
            }
        };
    }

    /// <summary>
    /// (Showing comment in Interface)
    /// Define Location of Merged Cells in the current Row. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Row. Notice that the Merged-Cells
    /// RowNumber should match with the Row RowNumber itself, otherwise an error will throw.
    /// </summary>
    public RowBuilder SetMergedCells(List<MergedBoundaryLocation> mergedCellsList)
    {
        if (mergedCellsList.Count > 0)
            CanBuild = true;

        Row.MergedCellsList = mergedCellsList;

        return this;
    }

    /// <summary>
    /// (Showing comment in Interface)
    /// In case we don't have any merge in the Row 
    /// </summary>
    /// <returns></returns>
    public RowBuilder NoMergedCells()
    {
        CanBuild = true;

        return this;
    }

    /// <summary>
    /// Set Row Styles including Bg, Font, Height, Borders and etc
    /// </summary>
    public RowBuilder SetStyle(RowStyle rowStyle)
    {
        Row.RowStyle = rowStyle;

        return this;
    }

    public Row Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Row because some necessary information not provided");

        return Row;
    }

}
