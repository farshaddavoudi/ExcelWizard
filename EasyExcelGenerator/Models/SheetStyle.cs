namespace EasyExcelGenerator.Models;

public class SheetStyle
{
    public SheetDirection? SheetDirection { get; set; } = null;

    public TextAlign? SheetDefaultTextAlign { get; set; } = null;

    /// <summary>
    /// Default column width for this worksheet.
    /// </summary>
    public double? SheetDefaultColumnWidth { get; set; } = null;

    /// <summary>
    /// Default row height for this worksheet.
    /// </summary>
    public double? SheetDefaultRowHeight { get; set; } = null;

    public SheetVisibility Visibility { get; set; } = SheetVisibility.Visible;
}