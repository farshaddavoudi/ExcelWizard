﻿using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using System;
using System.Collections.Generic;
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
    /// <param name="cellBuilder"> CellBuilder with Build() method at the end  </param>
    /// <param name="cellBuilders"> CellBuilder(s) with Build() method at the end of them </param>
    public static IExpectMergedCellsStatusRowBuilder SetCells(ICellBuilder cellBuilder, params ICellBuilder[] cellBuilders)
    {
        ICellBuilder[] cells = new[] { cellBuilder }.Concat(cellBuilders).ToArray();

        return new RowBuilder
        {
            Row = new Row
            {
                RowCells = cells.Select(c => (Cell)c).ToList()
            }
        };
    }

    /// <summary>
    /// Each Row contains one or more Cell(s). It is required as Row definition cannot be without Cells.
    /// </summary>
    /// <param name="cellBuilders"> CellBuilders with Build() method at the end of them </param>
    public static IExpectMergedCellsStatusRowBuilder SetCells(List<ICellBuilder> cellBuilders)
    {
        return new RowBuilder
        {
            Row = new Row
            {
                RowCells = cellBuilders.Select(c => (Cell)c).ToList()
            }
        };
    }

    /// <summary>
    /// (Showing comment in Interface)
    /// Define Location of Merged Cells in the current Row. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Row. Notice that the Merged-Cells
    /// RowNumber should match with the Row RowNumber itself, otherwise an error will throw.
    /// </summary>
    public IExpectStyleRowBuilder SetRowMergedCells(IMergeBuilder mergeBuilder, params IMergeBuilder[] mergeBuilders)
    {
        IMergeBuilder[] merges = new[] { mergeBuilder }.Concat(mergeBuilders).ToArray();

        CanBuild = true;

        Row.MergedCellsList = merges.Select(m => (MergedCells)m).ToList();

        return this;
    }

    public IExpectStyleRowBuilder SetRowMergedCells(List<IMergeBuilder> mergeBuilders)
    {
        CanBuild = true;

        Row.MergedCellsList = mergeBuilders.Select(m => (MergedCells)m).ToList();

        return this;
    }

    /// <summary>
    /// (Showing comment in Interface)
    /// In case we don't have any merge in the Row 
    /// </summary>
    /// <returns></returns>
    public IExpectStyleRowBuilder RowHasNoMerging()
    {
        CanBuild = true;

        return this;
    }

    public IExpectBuildMethodRowBuilder SetRowStyle(RowStyle rowStyle)
    {
        Row.RowStyle = rowStyle;

        return this;
    }

    public IExpectBuildMethodRowBuilder RowHasNoCustomStyle()
    {
        return this;
    }

    public IRowBuilder Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Row because some necessary information are not provided");

        return Row;
    }
}
