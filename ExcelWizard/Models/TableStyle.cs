using System.Drawing;

namespace ExcelWizard.Models;

public class TableStyle
{
    /// <summary>
    /// Set Background Color for entire Table. It will override the Sheet Background Color. Default inherit
    /// </summary>
    public Color? BackgroundColor { get; set; }

    /// <summary>
    /// Set the Font props (font-family, size, color, etc) for the entire table
    /// </summary>
    public TextFont? Font { get; set; }

    public TextAlign? TableTextAlign { get; set; }

    /// <summary>
    /// Set outside border of a table. Default is without border.
    /// </summary>
    public Border TableOutsideBorder { get; set; } = new() { BorderLineStyle = LineStyle.Thin, BorderColor = Color.LightGray };

    /// <summary>
    /// Set inside borders of table Cells. It do not effect the table Outside borders! Default is Thin border (Like Excel normal cells)
    /// </summary>
    public Border InsideCellsBorder { get; set; } = new() { BorderLineStyle = LineStyle.Thin, BorderColor = Color.LightGray };
}