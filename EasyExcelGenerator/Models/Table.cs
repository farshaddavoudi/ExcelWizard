
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace EasyExcelGenerator.Models;

public class Table : IValidatableObject
{
    public List<Row> TableRows { get; set; } = new();

    public CellLocation StartCellLocation
    {
        get
        {
            return TableRows.FirstOrDefault().StartCellLocation;
        }
    }  //TODO: Discuss with Shahab. The Rows has StartLocation itself, which one should be considered?
    //TODO: StartLocation and EndLocation for Table model are critical and should exist and be exact to create desired result

    public CellLocation EndLocation
    {
        get
        {
            return TableRows.LastOrDefault().EndCellLocation;
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
            var y = TableRows.LastOrDefault().EndCellLocation.Y - (TableRows.LastOrDefault().EndCellLocation.Y - TableRows.LastOrDefault().StartCellLocation.Y);
            return new CellLocation(TableRows.LastOrDefault().EndCellLocation.X + 1, y);
        }
    }

    public CellLocation NextVerticalCellLocation
    {
        get
        {
            var x = TableRows.LastOrDefault().EndCellLocation.X - (TableRows.LastOrDefault().EndCellLocation.X - TableRows.LastOrDefault().StartCellLocation.X);
            return new CellLocation(x, TableRows.LastOrDefault().EndCellLocation.Y + 1);
        }
    }

    public Cell GetCell(CellLocation cellLocation)
    {
        return TableRows[cellLocation.X - 1].Cells[cellLocation.Y - 1];
    }

    public List<Cell> GetCells(CellLocation startCellLocation, CellLocation endCellLocation)
    {
        List<Cell> cells = new();
        for (int i = startCellLocation.Y; i < endCellLocation.Y; i++)
        {
            cells.Add(GetCell(new CellLocation(startCellLocation.X, i)));
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