using ClosedXML.Report.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;

namespace EasyExcelGenerator.Models;

public class Row : IValidatableObject
{
    public List<Cell> Cells { get; set; } = new();


    // TODO: Move all Location relevant props to another class 

    public CellLocation StartCellLocation => Cells.First().CellLocation;

    public CellLocation EndCellLocation => Cells.Last().CellLocation;

    // TODO: Move all Style relevant props to another class

    public Color BackgroundColor { get; set; } = Color.White;

    public Color FontColor { get; set; } = Color.Black;

    // TODO: Add other props including: Bold, FontName, FontSize, Italic, Shadow, StrikeThrough

    public double? RowHeight { get; set; }

    public List<string> MergedCellsList { get; set; } = new();

    public Border InsideBorder { get; set; } = new();

    public Border OutsideBorder { get; set; } = new();

    public string? Formulas { get; set; }

    public CellLocation NextHorizontalCellLocation
    {
        get
        {
            var y = EndCellLocation.Y - (EndCellLocation.Y - StartCellLocation.Y);

            return new CellLocation(EndCellLocation!.X + 1, y);
        }
    }

    public CellLocation NextVerticalCellLocation
    {
        get
        {
            var x = EndCellLocation.X - (EndCellLocation.X - StartCellLocation.X); //TODO: ?? (x-(x-y) => answer always is y)

            return new CellLocation(x, EndCellLocation.Y + 1);
        }
    }

    public Cell AddCell()
    {
        Cell cell = new(NextHorizontalCellLocation);

        Cells.Add(cell);

        return cell;
    }

    public Cell? GetCell(int x)
    {
        return Cells.FirstOrDefault(c => c.CellLocation.X == x);
    }

    // Validations
    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        // TODO: Check below validations totally in phase 2

        if (Cells.Count == 0)
            yield return new ValidationResult("Row instance should contain one or more Cell(s)");

        // TODO: Check Y of StartLocation and EndLocation should be the equal and same with other Cells location Y property (Check with Shahab)

        // Checks Row cells have all Y location //TODO: Discuss with Shahab is it true validation or not
        if (Cells.Count != 0)
        {
            var firstCellYLoc = Cells.First().CellLocation.Y;

            foreach (var cell in Cells)
            {
                if (cell.CellLocation.Y != firstCellYLoc)
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
                if (c.ToString().IsNumber() && Cells.Count != 0 && Convert.ToInt32(c.ToString()) != Cells.First()?.CellLocation.Y)
                {
                    yield return new ValidationResult("In MergedCell Az:Bz the z should be the same with Row Y location");
                }
            }
        }
    }
}