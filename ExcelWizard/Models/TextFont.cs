using System.Drawing;

namespace ExcelWizard.Models;

public class TextFont
{
    public double? FontSize { get; set; }

    public bool? IsBold { get; set; }

    public Color? FontColor { get; set; } //Default null and get from Row in this case

    public string? FontName { get; set; }
}