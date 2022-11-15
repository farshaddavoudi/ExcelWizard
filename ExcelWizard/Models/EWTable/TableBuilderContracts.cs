using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;

namespace ExcelWizard.Models.EWTable;

public interface ITableBuilder
{

}

public interface IExpectRowsTableBuilder
{
    /// <summary>
    /// Each Table contains one or more Row(s). It is required as Table definition cannot be without Rows.
    /// </summary>
    IExpectMergedCellsStatusInManualProcessTableBuilder SetRows(params IRowBuilder[] tableRows);
}

public interface IExpectMergedCellsStatusInManualProcessTableBuilder
{
    /// <summary>
    /// Define of Merged Cells in the current Table. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Table. Notice that the Merged Cells
    /// should place into the Locations of the current Table, otherwise an error will throw.
    /// </summary>
    IExpectStyleTableBuilder SetTableMergedCells(params IMergeBuilder[] merges);

    /// <summary>
    /// In case we don't have any merge in the Table
    /// </summary>
    /// <returns></returns>
    IExpectStyleTableBuilder HasNoMergedCells();
}

public interface IExpectStyleTableBuilder
{
    /// <summary>
    /// Set Table Styles e.g. OutsideBorder, etc
    /// </summary>
    IExpectBuildMethodInManualTableBuilder SetStyle(TableStyle tableStyle);

    /// <summary>
    /// No custom styles for the table
    /// </summary>
    IExpectBuildMethodInManualTableBuilder NoCustomStyle();
}

public interface IExpectMergedCellsStatusInModelTableBuilder
{
    /// <summary>
    /// Define of Merged Cells in the current Table. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Table. Notice that the Merged Cells
    /// should place into the Locations of the current Table, otherwise an error will throw.
    /// </summary>
    IExpectBuildMethodInModelTableBuilder SetMergedCells(params IMergeBuilder[] merges);

    /// <summary>
    /// In case we don't have any merge in the Table
    /// </summary>
    /// <returns></returns>
    IExpectBuildMethodInModelTableBuilder NoMergedCells();
}

public interface IExpectBuildMethodInModelTableBuilder
{
    Table Build();
}

public interface IExpectBuildMethodInManualTableBuilder
{
    Table Build();
}