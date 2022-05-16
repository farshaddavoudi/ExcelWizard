using System.Drawing;

namespace ExcelWizard.Models;

public class CellStyle
{
    public TextFont CellFont { get; set; } = new();

    /// <summary>
    /// Set Wordwrap for the Cell content. Default is false
    /// </summary>
    public bool Wordwrap { get; set; }

    public TextAlign? CellTextAlign { get; set; }

    /// <summary>
    /// Set Background Color for Cell. It will override the Table or Row or Column Background Colors. Default inherit
    /// </summary>
    public Color? BackgroundColor { get; set; }
}