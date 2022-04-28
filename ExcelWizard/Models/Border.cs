using System.Drawing;

namespace ExcelWizard.Models;

public class Border
{
    public LineStyle BorderLineStyle { get; set; } = LineStyle.None;

    public Color BorderColor { get; set; } = Color.Black;
}