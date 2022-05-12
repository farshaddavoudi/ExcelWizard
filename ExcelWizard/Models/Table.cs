
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelWizard.Models;

public class Table : IValidatableObject
{
    public List<Row> TableRows { get; set; } = new();

    public CellLocation? StartCellLocation => TableRows.FirstOrDefault()?.GetRowStartCellLocation(); //TODO: Discuss with Shahab. The Rows has StartLocation itself, which one should be considered?
    // TODO: StartLocation and EndLocation for Table model are critical and should exist and be exact to create desired result
    // TODO: Remove StartLoc and EndLoc. It should calculated by Cells Loc

    public CellLocation EndLocation
    {
        get
        {
            return TableRows.LastOrDefault().GetRowEndCellLocation();
        }

    } //TODO: above question

    public Border InlineBorder { get; set; } = new();//TODO: What it is? Inside border can be set on cells or columns or rows

    public Border OutsideBorder { get; set; } = new();

    public bool IsBordered { get; set; } //TODO? What is this? isn't it the default one?

    public List<string> MergedCells { get; set; } = new();

    public int RowsCount => TableRows.Count;

    public CellLocation NextHorizontalCellLocation
    {
        get
        {
            var y = TableRows.LastOrDefault().GetRowEndCellLocation().RowNumber - (TableRows.LastOrDefault().GetRowEndCellLocation().RowNumber - TableRows.LastOrDefault().GetRowStartCellLocation().RowNumber);
            return new CellLocation(TableRows.LastOrDefault().GetRowEndCellLocation().ColumnNumber + 1, y);
        }
    }

    public CellLocation NextVerticalCellLocation
    {
        get
        {
            var x = TableRows.LastOrDefault().GetRowEndCellLocation().ColumnNumber - (TableRows.LastOrDefault().GetRowEndCellLocation().ColumnNumber - TableRows.LastOrDefault().GetRowStartCellLocation().ColumnNumber);
            return new CellLocation(x, TableRows.LastOrDefault().GetRowEndCellLocation().RowNumber + 1);
        }
    }

    public Cell GetCell(CellLocation cellLocation)
    {
        return TableRows[cellLocation.ColumnNumber - 1].Cells[cellLocation.RowNumber - 1];
    }

    public List<Cell> GetCells(CellLocation startCellLocation, CellLocation endCellLocation)
    {
        List<Cell> cells = new();
        for (int i = startCellLocation.RowNumber; i < endCellLocation.RowNumber; i++)
        {
            cells.Add(GetCell(new CellLocation(startCellLocation.ColumnNumber, i)));
        }

        return cells;
    }

    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        if (false)
            yield return new ValidationResult("");
        // TODO: Discuess with Shahab. Shouldn't Rows in a Table have common features like Same StartLocation.X and things like
    }
}