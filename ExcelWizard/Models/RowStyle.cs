using System.Drawing;

namespace ExcelWizard.Models;

public class RowStyle
{
    /// <summary>
    /// Set Background Color for entire Row. It will override the Table Background Colors. Default inherit
    /// </summary>
    public Color? BackgroundColor { get; set; }

    public TextFont Font { get; set; } = new();

    public double? RowHeight { get; set; }

    public Border InsideBorder { get; set; } = new();

    public Border OutsideBorder { get; set; } = new();
}