using System.Drawing;

namespace ExcelWizard.Models;

public class TextFont
{
    public double? FontSize { get; set; }

    public bool? IsBold { get; set; }

    public Color? FontColor { get; set; }

    public string? FontName { get; set; }
}