using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelWizard.Models;

public class Row : IValidatableObject
{
    // Props

    /// <summary>
    /// Each Row contains one or more Cell(s). It is required as Row definition cannot be without Cells.
    /// </summary>
    public List<Cell> Cells { get; set; } = new();

    /// <summary>
    /// Set Row Styles including Bg, Font, Height, Borders and etc
    /// </summary>
    public RowStyle RowStyle { get; set; } = new();

    /// <summary>
    /// Arbitrary property to define Merged Cells in the current Row. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Row. Notice that the Merged-Cells
    /// RowNumber should match with the Row RowNumber itself, otherwise an error will throw.
    /// </summary>
    public List<MergeCells> MergedCellsList { get; set; } = new();

    // TODO: R&D about it
    public string? Formulas { get; set; }

    // Methods

    /// <summary>
    /// Get current Row Starting Cell Automatically by its Cells Location
    /// </summary>
    /// <returns></returns>
    public CellLocation GetRowStartCellLocation()
    {
        var rowColumns = Cells.Select(c => c.CellLocation.ColumnNumber).ToList();

        var rowNumber = Cells.First().CellLocation.RowNumber; //All Cells in a Row have equal RowNumber

        return new CellLocation(rowColumns.Min(), rowNumber);
    }

    /// <summary>
    /// Get current Row Ending Cell Automatically by its Cells Location
    /// </summary>
    /// <returns></returns>
    public CellLocation GetRowEndCellLocation()
    {
        var rowColumns = Cells.Select(c => c.CellLocation.ColumnNumber).ToList();

        var rowNumber = Cells.First().CellLocation.RowNumber; //All Cells in a Row have equal RowNumber

        return new CellLocation(rowColumns.Max(), rowNumber);
    }

    public CellLocation GetNextHorizontalCellLocationAfterRow()
    {
        var rowEndLocation = GetRowEndCellLocation();

        return new CellLocation(rowEndLocation.ColumnNumber + 1, rowEndLocation.RowNumber);
    }

    public Cell? GetRowCellByColumnNumber(int columnNumber)
    {
        return Cells.FirstOrDefault(c => c.CellLocation.ColumnNumber == columnNumber);
    }

    // Validations
    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        // Row definition cannot have no Cells
        if (Cells.Count == 0)
            yield return new ValidationResult("Row instance should contain one or more Cell(s)");

        // Check Y of StartLocation and EndLocation should be the equal and same with other Cells location Y property (Check with Shahab)
        if (Cells.Select(c => c.CellLocation.RowNumber).Distinct().ToList().Count != 1)
            yield return new ValidationResult("Invalid Row. All Row Cells should have equal RowNumber!");

        // Check MergedCells format
        var currentRowNumber = Cells.First().CellLocation.RowNumber;

        foreach (var cellsToMerge in MergedCellsList)
        {
            if (cellsToMerge.FirstCellLocation is null || cellsToMerge.LastCellLocation is null)
                yield return new ValidationResult("Something is not right about MergedCells format. FirstCellLocation and LastCellLocations cannot be null");

            if (cellsToMerge.FirstCellLocation!.RowNumber != currentRowNumber)
                yield return new ValidationResult("Something is not right about MergedCells format. The RowNumber of FirstCellLocation should be equal with the Row RowNumber!");

            if (cellsToMerge.LastCellLocation!.RowNumber != currentRowNumber)
                yield return new ValidationResult("Something is not right about MergedCells format. The RowNumber of LastCellLocation should be equal with the Row RowNumber!");
        }
    }
}