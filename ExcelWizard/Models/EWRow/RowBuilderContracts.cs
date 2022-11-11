using System.Collections.Generic;

namespace ExcelWizard.Models.EWRow;

public interface IRowBuilder
{

}

public interface IExpectMergedCellsStatusRowBuilder
{
    /// <summary>
    /// Define Location of Merged Cells in the current Row. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Row. Notice that the Merged-Cells
    /// RowNumber should match with the Row RowNumber itself, otherwise an error will throw.
    /// </summary>
    RowBuilder SetMergedCells(List<MergedBoundaryLocation> mergedCellsList);

    /// <summary>
    /// In case we don't have any merge in the Row
    /// </summary>
    /// <returns></returns>
    RowBuilder NoMergedCells();
}