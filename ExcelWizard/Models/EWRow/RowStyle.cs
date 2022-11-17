using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace ExcelWizard.Models;

public class RowStyle
{
    /// <summary>
    /// Set Background Color for entire Row. It will override the Table Background Colors. Default inherit
    /// </summary>
    public Color? BackgroundColor { get; set; }

    /// <summary>
    /// Set Font for entire Row. It will override the Table Font. Default inherit
    /// </summary>
    public TextFont? Font { get; set; }

    public TextAlign? RowTextAlign { get; set; }

    public double? RowHeight { get; set; }

    public Border InsideCellsBorder { get; set; } = new();

    public Border RowOutsideBorder { get; set; } = new();
}