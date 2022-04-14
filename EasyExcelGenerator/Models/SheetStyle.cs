namespace EasyExcelGenerator.Models;

public class SheetStyle
{
    public SheetDirection? SheetDirection { get; set; } = null;

    public TextAlign? DefaultTextAlign { get; set; } = null;

    /// <summary>
    /// Default column width for this worksheet.
    /// </summary>
    public double? DefaultColumnWidth { get; set; } = null;

    /// <summary>
    /// Default row height for this worksheet.
    /// </summary>
    public double? DefaultRowHeight { get; set; } = null;

    public SheetVisibility Visibility { get; set; } = SheetVisibility.Visible;
}