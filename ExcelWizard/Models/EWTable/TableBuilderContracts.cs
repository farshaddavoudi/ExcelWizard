using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;

namespace ExcelWizard.Models.EWTable;

public interface ITableBuilder
{
    /// <summary>
    /// Get table Starting Cell Automatically
    /// </summary>
    /// <param name="considerMergedCells"> Sometimes Merging Cells definition cause the Table goes beyond boundary of its Cells which is normal. The table actually finish when its Merged Cells finishes </param>
    /// <returns></returns>
    CellLocation GetTableFirstCellLocation(bool considerMergedCells = true);

    /// <summary>
    /// Get table Ending Cell Automatically
    /// </summary>
    /// <param name="considerMergedCells"> Sometimes Merging Cells definition cause the Table goes beyond boundary of its Cells which is normal. The table actually finish when its Merged Cells finishes </param>
    /// <returns></returns>
    CellLocation GetTableLastCellLocation(bool considerMergedCells = true);

    int GetNextHorizontalColumnNumberAfterTable();
    int GetNextVerticalRowNumberAfterTable();

    /// <summary>
    ///  Get the Table Cell by its location. The Location should be inside Table territory
    /// </summary>
    Cell? GetTableCell(CellLocation cellLocation);

    void ValidateTableInstance();
}

public interface IExpectRowsTableBuilder
{
    /// <summary>
    /// Each Table contains one or more Row(s). It is required as Table definition cannot be without Rows.
    /// </summary>
    /// <param name="rowBuilders"> RowBuilder(s) with Build() method at the end of them </param>
    IExpectMergedCellsStatusInManualProcessTableBuilder SetRows(params IRowBuilder[] rowBuilders);
}

public interface IExpectMergedCellsStatusInManualProcessTableBuilder
{
    /// <summary>
    /// Define of Merged Cells in the current Table. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Table. Notice that the Merged Cells
    /// should place into the Locations of the current Table, otherwise an error will throw.
    /// </summary>
    /// <param name="mergeBuilders"> MergeBuilder(s) with Build() method at the end of them </param>
    IExpectStyleTableBuilder SetTableMergedCells(params IMergeBuilder[] mergeBuilders);

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
    /// <param name="mergeBuilders"> MergeBuilder(s) with Build() method at the end of them </param>
    IExpectBuildMethodInModelTableBuilder SetMergedCells(params IMergeBuilder[] mergeBuilders);

    /// <summary>
    /// In case we don't have any merge in the Table
    /// </summary>
    /// <returns></returns>
    IExpectBuildMethodInModelTableBuilder NoMergedCells();
}

public interface IExpectBuildMethodInModelTableBuilder
{
    ITableBuilder Build();
}

public interface IExpectBuildMethodInManualTableBuilder
{
    ITableBuilder Build();
}