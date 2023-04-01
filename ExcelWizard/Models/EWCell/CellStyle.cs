using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace ExcelWizard.Models;

public class CellStyle
{
    /// <summary>
    /// Set Font for the Cell. It will override the Table and Row Fonts. Default inherit
    /// </summary>
    public TextFont? Font { get; set; }

    /// <summary>
    /// Set Wordwrap for the Cell content. Default is false
    /// </summary>
    public bool Wordwrap { get; set; }

    public TextAlign? CellTextAlign { get; set; }

    /// <summary>
    /// Set Background Color for Cell. It will override the Table or Row or Column Background Colors. Default inherit
    /// </summary>
    public Color? BackgroundColor { get; set; }

    /// <summary>
    /// Set outside border of a table. Default is without border.
    /// </summary>
    public Border? CellBorder { get; set; }
}