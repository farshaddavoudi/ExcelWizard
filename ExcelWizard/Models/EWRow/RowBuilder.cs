using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using System;
using System.Linq;

namespace ExcelWizard.Models.EWRow;

public class RowBuilder : IExpectMergedCellsStatusRowBuilder, IExpectBuildMethodRowBuilder, IExpectStyleRowBuilder
{
    private RowBuilder() { }

    private Row Row { get; set; } = new();
    private bool CanBuild { get; set; }

    /// <summary>
    /// Each Row contains one or more Cell(s). It is required as Row definition cannot be without Cells.
    /// </summary>
    public static IExpectMergedCellsStatusRowBuilder SetCells(params ICellBuilder[] rowCells)
    {
        return new RowBuilder
        {
            Row = new Row
            {
                RowCells = rowCells.Select(c => (Cell)c).ToList()
            }
        };
    }

    /// <summary>
    /// (Showing comment in Interface)
    /// Define Location of Merged Cells in the current Row. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Row. Notice that the Merged-Cells
    /// RowNumber should match with the Row RowNumber itself, otherwise an error will throw.
    /// </summary>
    public IExpectStyleRowBuilder SetRowMergedCells(params IMergeBuilder[] merges)
    {
        if (merges.Length > 0)
            CanBuild = true;

        Row.MergedCellsList = merges.Select(m => (MergedCells)m).ToList();

        return this;
    }

    /// <summary>
    /// (Showing comment in Interface)
    /// In case we don't have any merge in the Row 
    /// </summary>
    /// <returns></returns>
    public IExpectStyleRowBuilder NoMergedCells()
    {
        CanBuild = true;

        return this;
    }

    public IExpectBuildMethodRowBuilder SetStyle(RowStyle rowStyle)
    {
        Row.RowStyle = rowStyle;

        return this;
    }

    public IExpectBuildMethodRowBuilder NoCustomStyle()
    {
        return this;
    }

    public Row Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Row because some necessary information are not provided");

        return Row;
    }
}
