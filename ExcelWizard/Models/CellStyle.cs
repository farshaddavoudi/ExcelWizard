namespace ExcelWizard.Models;

public class CellStyle
{
    public TextFont CellFont { get; set; } = new();

    /// <summary>
    /// Set Wordwrap for the Cell content. Default is false
    /// </summary>
    public bool Wordwrap { get; set; }

    public TextAlign? CellTextAlign { get; set; }
}