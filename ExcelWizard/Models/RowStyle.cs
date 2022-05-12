using System.Drawing;

namespace ExcelWizard.Models;

public class RowStyle
{
    public Color BackgroundColor { get; set; } = Color.White;

    public TextFont Font { get; set; } = new();

    public double? RowHeight { get; set; }

    public Border InsideBorder { get; set; } = new();

    public Border OutsideBorder { get; set; } = new();
}