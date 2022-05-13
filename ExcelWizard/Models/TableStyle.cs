using System.Drawing;

namespace ExcelWizard.Models;

public class TableStyle
{
    /// <summary>
    /// Set outside border of a table. Default is without border.
    /// </summary>
    public Border OutsideBorder { get; set; } = new();

    /// <summary>
    /// Set inline or inside border of table Cells. Default is Thin border (Like Excel normal cells)
    /// </summary>
    public Border CellsSeparatorBorder { get; set; } = new() { BorderLineStyle = LineStyle.Thin, BorderColor = Color.LightGray };
}