using ExcelWizard.Models.EWMerge;

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
    IExpectStyleRowBuilder SetRowMergedCells(params MergedCells[] mergedCells);

    /// <summary>
    /// In case we don't have any merge in the Row
    /// </summary>
    /// <returns></returns>
    IExpectStyleRowBuilder NoMergedCells();
}

public interface IExpectStyleRowBuilder
{
    /// <summary>
    /// Set Row Styles including Bg, Font, Height, Borders and etc
    /// </summary>
    IExpectBuildMethodRowBuilder SetStyle(RowStyle rowStyle);

    /// <summary>
    /// No custom styles for the row
    /// </summary>
    IExpectBuildMethodRowBuilder NoCustomStyle();
}

public interface IExpectBuildMethodRowBuilder
{
    Row Build();
}