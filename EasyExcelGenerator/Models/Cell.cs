namespace EasyExcelGenerator.Models;

public class Cell
{
    public Cell(CellLocation cellLocation)
    {
        CellLocation = cellLocation;
    }

    public string? Name { get; set; } //TODO: Add Name property somehow as column (cell) identifier

    public object Value { get; set; }

    public CellLocation CellLocation { get; set; }

    public TextFont Font { get; set; } = new();

    public bool Wordwrap { get; set; }

    public TextAlign? CellTextAlign { get; set; }

    public CellType CellType { get; set; } = CellType.General;

    public bool IsCellVisible { get; set; } = true;

    public bool? IsCellLocked { get; set; } = null; //Default is null, and it gets Sheet "IsLocked" property value in this case, but if specified, it will override it
}