using System.Collections.Generic;

namespace EasyExcelGenerator.Models;

public class GridLayoutExcelBuilder
{
    public string? GeneratedFileName { get; set; }

    public List<GridExcelSheet> Sheets { get; set; } = new();
}