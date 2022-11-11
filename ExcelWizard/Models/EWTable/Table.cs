using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWRow;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelWizard.Models.EWTable;

public class Table
{
    // Props

    /// <summary>
    /// Each Table contains one or more Row(s). It is required as Table definition cannot be without Rows.
    /// </summary>
    public List<Row2> TableRows { get; set; } = new();

    /// <summary>
    /// Set Table Styles e.g. OutsideBorder, etc
    /// </summary>
    public TableStyle TableStyle { get; set; } = new();

    /// <summary>
    /// Arbitrary property to define of Merged Cells in the current Table. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Table. Notice that the Merged Cells
    /// should place into the Locations of the current Table, otherwise an error will throw.
    /// </summary>
    public List<MergedCells> MergedCellsList { get; set; } = new();

    // Methods

    /// <summary>
    /// Get table Starting Cell Automatically
    /// </summary>
    /// <param name="considerMergedCells"> Sometimes Merging Cells definition cause the Table goes beyond boundary of its Cells which is normal. The table actually finish when its Merged Cells finishes </param>
    /// <returns></returns>
    public CellLocation GetTableFirstCellLocation(bool considerMergedCells = true)
    {
        var tableRowNumbers = TableRows.Select(r => r.GetRowNumber()).ToList();

        var startCellRowNumber = tableRowNumbers.Min();

        var firstCellLocationFromTableWithoutConsideringMergedCells = TableRows.First(r => r.GetRowNumber() == startCellRowNumber).GetRowFirstCellLocation();

        if (considerMergedCells is false || MergedCellsList.Count == 0)
            return firstCellLocationFromTableWithoutConsideringMergedCells;

        // Order priority is based on RowNumber rather than ColumnNumber

        var firstRowOfMergedCells = MergedCellsList.Select(mc => mc.MergedBoundaryLocation.FirstCellLocation!.RowNumber).ToList().Min();

        if (firstRowOfMergedCells >= firstCellLocationFromTableWithoutConsideringMergedCells.RowNumber)
            return firstCellLocationFromTableWithoutConsideringMergedCells;

        // Now we know about our target location RowNumber. Lets find its ColumnNumber
        // Notice: Cell does not exist in Table location territory
        var firstColumnNumberOfMergedCells = MergedCellsList
            .Where(mc => mc.MergedBoundaryLocation.FirstCellLocation!.RowNumber == firstRowOfMergedCells)
            .Select(mc => mc.MergedBoundaryLocation.FirstCellLocation!.ColumnNumber)
            .Min();

        return new CellLocation(firstColumnNumberOfMergedCells, firstRowOfMergedCells);
    }

    /// <summary>
    /// Get table Ending Cell Automatically
    /// </summary>
    /// <param name="considerMergedCells"> Sometimes Merging Cells definition cause the Table goes beyond boundary of its Cells which is normal. The table actually finish when its Merged Cells finishes </param>
    /// <returns></returns>
    public CellLocation GetTableLastCellLocation(bool considerMergedCells = true)
    {
        var tableRowNumbers = TableRows.Select(r => r.GetRowNumber()).ToList();

        var endCellRowNumber = tableRowNumbers.Max();

        var lastCellLocationFromTableWithoutConsideringMergedCells = TableRows.First(r => r.GetRowNumber() == endCellRowNumber).GetRowLastCellLocation();

        if (considerMergedCells is false || MergedCellsList.Count == 0)
            return lastCellLocationFromTableWithoutConsideringMergedCells;

        // Order priority is based on RowNumber rather than ColumnNumber

        var lastRowOfMergedCells = MergedCellsList.Select(mc => mc.MergedBoundaryLocation.LastCellLocation!.RowNumber).ToList().Max();

        if (lastRowOfMergedCells <= lastCellLocationFromTableWithoutConsideringMergedCells.RowNumber)
            return lastCellLocationFromTableWithoutConsideringMergedCells;

        // Now we know about our target location RowNumber. Lets find its ColumnNumber
        // Notice: Cell does not exist in Table location territory
        var lastColumnNumberOfMergedCells = MergedCellsList
            .Where(mc => mc.MergedBoundaryLocation.LastCellLocation!.RowNumber == lastRowOfMergedCells)
            .Select(mc => mc.MergedBoundaryLocation.LastCellLocation!.ColumnNumber)
            .Max();

        return new CellLocation(lastColumnNumberOfMergedCells, lastRowOfMergedCells);
    }

    public int GetNextHorizontalColumnNumberAfterTable()
    {
        var lastTableCell = GetTableLastCellLocation();

        return lastTableCell.ColumnNumber + 1;
    }

    public int GetNextVerticalRowNumberAfterTable()
    {
        var lastTableCell = GetTableLastCellLocation();

        return lastTableCell.RowNumber + 1;
    }

    /// <summary>
    ///  Get the Table Cell by its location. The Location should be inside Table territory
    /// </summary>
    public Cell? GetTableCell(CellLocation cellLocation)
    {
        var cellRow = TableRows.FirstOrDefault(r => r.GetRowNumber() == cellLocation.RowNumber);

        return cellRow?.GetRowCellByColumnNumber(cellLocation.ColumnNumber);
    }

    // Validations
    public void ValidateTableInstance()
    {
        // Table definition cannot have no Rows
        if (TableRows.Count == 0)
            throw new ValidationException("Table instance should contain one or more Row(s)");

        // Check Providing MergedCells Items
        if (MergedCellsList.Count != 0)
        {
            if (MergedCellsList.Any(mc =>
                    mc.MergedBoundaryLocation.FirstCellLocation is null ||
                    mc.MergedBoundaryLocation.LastCellLocation is null))
            {
                throw new ValidationException("Table Merged Cells start and end locations are required");
            }
        }

        // Check Merged Cells
        foreach (var cellsToMerge in MergedCellsList)
        {
            if (cellsToMerge.MergedBoundaryLocation.FirstCellLocation is null || cellsToMerge.MergedBoundaryLocation.LastCellLocation is null)
                throw new ValidationException("Something is not right about MergedCells format. FirstCellLocation and LastCellLocations cannot be null");

            if (cellsToMerge.MergedBoundaryLocation.FirstCellLocation!.RowNumber < GetTableFirstCellLocation().RowNumber)
                throw new ValidationException("Something is not right about MergedCells format. The RowNumber of FirstCellLocation cannot be little than TableFirstCellLocation");

            if (cellsToMerge.MergedBoundaryLocation.FirstCellLocation!.ColumnNumber < GetTableFirstCellLocation().ColumnNumber)
                throw new ValidationException("Something is not right about MergedCells format. The ColumnNumber of FirstCellLocation cannot be little than TableFirstCellLocation");

            if (cellsToMerge.MergedBoundaryLocation.LastCellLocation!.RowNumber > GetTableLastCellLocation().RowNumber)
                throw new ValidationException("Something is not right about MergedCells format. The RowNumber of LastCellLocation cannot be greater than TableLastCellLocation");

            if (cellsToMerge.MergedBoundaryLocation.LastCellLocation!.ColumnNumber > GetTableLastCellLocation().ColumnNumber)
                throw new ValidationException("Something is not right about MergedCells format. The ColumnNumber of LastCellLocation cannot be greater than TableLastCellLocation");
        }

        // Not repetitive Locations
        var isAllRowsUnique = TableRows.Select(r => r.GetRowNumber()).Distinct().Count() == TableRows.Count;

        if (isAllRowsUnique is false)
            throw new ValidationException("There are some repetitive rows in the Table. Make all rows unique");

    }
}