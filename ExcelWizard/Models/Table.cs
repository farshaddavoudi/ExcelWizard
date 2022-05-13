
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelWizard.Models;

public class Table : IValidatableObject
{
    // Props

    /// <summary>
    /// Each Table contains one or more Row(s). It is required as Table definition cannot be without Rows.
    /// </summary>
    public List<Row> TableRows { get; set; } = new();

    /// <summary>
    /// Set Table Styles e.g. OutsideBorder, etc
    /// </summary>
    public TableStyle TableStyle { get; set; } = new();

    /// <summary>
    /// Arbitrary property to define of Merged Cells in the current Table. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Table. Notice that the Merged Cells
    /// should place into the Locations of the current Table, otherwise an error will throw.
    /// </summary>
    public List<MergeStartEndLocation> MergedCellsList { get; set; } = new();

    // Methods

    /// <summary>
    /// Get table Starting Cell Automatically
    /// </summary>
    /// <returns></returns>
    public CellLocation GetTableFirstCellLocation()
    {
        var tableRowNumbers = TableRows.Select(r => r.GetRowNumber()).ToList();

        var startCellRowNumber = tableRowNumbers.Min();

        return TableRows.First(r => r.GetRowNumber() == startCellRowNumber).GetRowFirstCellLocation();
    }

    /// <summary>
    /// Get table Ending Cell Automatically
    /// </summary>
    /// <returns></returns>
    public CellLocation GetTableLastCellLocation()
    {
        var tableRowNumbers = TableRows.Select(r => r.GetRowNumber()).ToList();

        var endCellRowNumber = tableRowNumbers.Max();

        return TableRows.First(r => r.GetRowNumber() == endCellRowNumber).GetRowLastCellLocation();
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

    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        // Table definition cannot have no Rows
        if (TableRows.Count == 0)
            yield return new ValidationResult("Table instance should contain one or more Row(s)");

        // Check Merged Cells
        foreach (var cellsToMerge in MergedCellsList)
        {
            if (cellsToMerge.FirstCellLocation is null || cellsToMerge.LastCellLocation is null)
                yield return new ValidationResult("Something is not right about MergedCells format. FirstCellLocation and LastCellLocations cannot be null");

            if (cellsToMerge.FirstCellLocation!.RowNumber < GetTableFirstCellLocation().RowNumber)
                yield return new ValidationResult("Something is not right about MergedCells format. The RowNumber of FirstCellLocation cannot be little than TableFirstCellLocation");

            if (cellsToMerge.FirstCellLocation!.ColumnNumber < GetTableFirstCellLocation().ColumnNumber)
                yield return new ValidationResult("Something is not right about MergedCells format. The ColumnNumber of FirstCellLocation cannot be little than TableFirstCellLocation");

            if (cellsToMerge.LastCellLocation!.RowNumber > GetTableLastCellLocation().RowNumber)
                yield return new ValidationResult("Something is not right about MergedCells format. The RowNumber of LastCellLocation cannot be greater than TableLastCellLocation");
            
            if (cellsToMerge.LastCellLocation!.ColumnNumber > GetTableLastCellLocation().ColumnNumber)
                yield return new ValidationResult("Something is not right about MergedCells format. The ColumnNumber of LastCellLocation cannot be greater than TableLastCellLocation");
        }

        // Not repetitive Locations
        var isAllRowsUnique = TableRows.Select(r => r.GetRowNumber()).Distinct().Count() == TableRows.Count;

        if (isAllRowsUnique is false)
            yield return new ValidationResult("There are some repetitive rows in the Table. Make all rows unique");
    }
}