using System.Collections.Generic;

namespace ExcelWizard.Models.EWGridLayout;

public class GridLayoutExcelModel
{
    public string? GeneratedFileName { get; set; }

    public List<GridExcelSheet> Sheets { get; set; } = new();
}