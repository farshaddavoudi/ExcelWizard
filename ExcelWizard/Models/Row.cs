using ClosedXML.Report.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelWizard.Models;

public class Row : IValidatableObject
{
    /// <summary>
    /// Each Row contains one or more Cell(s). It is required as Row definition cannot be without Cells.
    /// </summary>
    public List<Cell> Cells { get; set; } = new();

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

    /// <summary>
    /// Set Row Styles including Bg, Font, Height, Borders and etc
    /// </summary>
    public RowStyle RowStyle { get; set; } = new();

    // TODO: Comment on it and R&D about it
    public List<string> MergedCellsList { get; set; } = new();

    // TODO: R&D about it
    public string? Formulas { get; set; }

    // Methods
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

        // TODO: Check Y of StartLocation and EndLocation should be the equal and same with other Cells location Y property (Check with Shahab)

        // Checks Row cells have all Y location //TODO: Discuss with Shahab is it true validation or not
        if (Cells.Count != 0)
        {
            var firstCellYLoc = Cells.First().CellLocation.RowNumber;

            foreach (var cell in Cells)
            {
                if (cell.CellLocation.RowNumber != firstCellYLoc)
                    yield return new ValidationResult("All row cells should have same Y location");
            }
        }

        // Check MergedCells format
        foreach (var cellsToMerge in MergedCellsList)
        {
            if (string.IsNullOrWhiteSpace(cellsToMerge) || cellsToMerge.Contains(":") is false)
                yield return
                    new ValidationResult("Something is not right about MergedCells format specified in Row model");

            // A2:B2 should be along with cells with locationY=2 //TODO: Confirm it with Shahab
            foreach (var c in cellsToMerge!.ToCharArray())
            {
                if (c.ToString().IsNumber() && Cells.Count != 0 && Convert.ToInt32(c.ToString()) != Cells.First()?.CellLocation.RowNumber)
                {
                    yield return new ValidationResult("In MergedCell Az:Bz the z should be the same with Row Y location");
                }
            }
        }
    }
}