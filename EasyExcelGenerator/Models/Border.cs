using System.Drawing;

namespace EasyExcelGenerator.Models;

public class Border
{
    public LineStyle BorderLineStyle { get; set; } = LineStyle.None;

    public Color BorderColor { get; set; } = Color.Black;
}