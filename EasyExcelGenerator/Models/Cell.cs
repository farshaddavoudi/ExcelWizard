using System;

namespace EasyExcelGenerator.Models;

public class Cell
{
    public Cell(CellLocation cellLocation)
    {
        CellLocation = cellLocation;
    }

    public string? Name { get; set; } //TODO: Add Name property somehow as column (cell) identifier

    internal Type? Type { get; set; }

    public object Value { get; set; }

    public CellLocation CellLocation { get; set; }

    public bool Wordwrap { get; set; }

    public TextAlign? TextAlign { get; set; }

    public CellType CellType { get; set; } = CellType.General;

    public bool Visible { get; set; } = true;

    public bool? IsLocked { get; set; } = null; //Default is null, and it gets Sheet "IsLocked" property value in this case, but if specified, it will override it
}