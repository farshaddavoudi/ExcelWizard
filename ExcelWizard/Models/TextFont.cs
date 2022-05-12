using System.Drawing;

namespace ExcelWizard.Models;

public class TextFont
{
    public double? FontSize { get; set; }

    public bool? IsBold { get; set; }

    /// <summary>
    /// Override the Row Font color. Default is inherit
    /// </summary>
    public Color? FontColor { get; set; }

    public string? FontName { get; set; }
}