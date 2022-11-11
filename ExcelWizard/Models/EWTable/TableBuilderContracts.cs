using ExcelWizard.Models.EWRow;
using System.Collections.Generic;

namespace ExcelWizard.Models.EWTable;

public interface ITableBuilder
{

}

public interface IExpectRowsTableBuilder
{
    /// <summary>
    /// Each Table contains one or more Row(s). It is required as Table definition cannot be without Rows.
    /// </summary>
    IExpectMergedCellsStatusTableBuilder SetRows(List<Row> tableRows);
}

public interface IExpectMergedCellsStatusTableBuilder
{
    /// <summary>
    /// Define of Merged Cells in the current Table. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Table. Notice that the Merged Cells
    /// should place into the Locations of the current Table, otherwise an error will throw.
    /// </summary>
    IExpectStyleTableBuilder SetMergedCells(List<MergedCells> mergedCellsList);

    /// <summary>
    /// In case we don't have any merge in the Table
    /// </summary>
    /// <returns></returns>
    IExpectStyleTableBuilder NoMergedCells();
}

public interface IExpectStyleTableBuilder
{
    /// <summary>
    /// Set Table Styles e.g. OutsideBorder, etc
    /// </summary>
    TableBuilder SetStyle(TableStyle tableStyle);

    /// <summary>
    /// No custom styles for the table
    /// </summary>
    TableBuilder NoCustomStyle();
}