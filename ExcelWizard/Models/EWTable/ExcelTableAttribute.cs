using ExcelWizard.Models.EWStyles;
using System;
using System.Drawing;

namespace ExcelWizard.Models.EWTable;

/// <summary>
/// Configure the Table generic properties in a complex Excel schema
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ExcelTableAttribute : Attribute
{
    /// <summary>
    /// whether the table renders with header or without it
    /// </summary>
    public bool HasHeader { get; set; } = true;

    /// <summary>
    /// How many rows does the Header occupy. It will be merged automatically
    /// </summary>
    public int HeaderOccupyingRowsNo { get; set; } = 1;

    /// <summary>
    /// Table header background color. Will be ignored if HasHeader is set to false
    /// </summary>
    public KnownColor HeaderBackgroundColor { get; set; } = KnownColor.Transparent;

    /// <summary>
    /// Table Cells TextAlign
    /// </summary>
    public TextAlign TextAlign { get; set; } = TextAlign.Inherit;

    /// <summary>
    /// Table Cells FontFamily name
    /// </summary>
    public string? FontName { get; set; }

    /// <summary>
    /// Table Cells FontColor
    /// </summary>
    public KnownColor FontColor { get; set; } = KnownColor.Black;

    /// <summary>
    /// Table Cells FontSize. If 0 it means default FontSize
    /// </summary>
    public int FontSize { get; set; }

    /// <summary>
    /// Font Weight for entire Table. Inherit is equal to default here which is bold for Header and normal for Cells
    /// </summary>
    public FontWeight FontWeight { get; set; } = FontWeight.Inherit;

    /// <summary>
    /// Table all data Cells background 
    /// </summary>
    public KnownColor DataBackgroundColor { get; set; } = KnownColor.Transparent;

    /// <summary>
    /// Set outside border of a table. Default is thin border.
    /// </summary>
    public LineStyle OutsideBorderStyle { get; set; } = LineStyle.Thin;

    /// <summary>
    /// Set outside border of a table. Default is color border.
    /// </summary>
    public KnownColor OutsideBorderColor { get; set; } = KnownColor.LightGray;

    /// <summary>
    /// Set style of inside borders of table Cells. It do not effect the table Outside borders! Default is Thin border (Like Excel normal cells)
    /// </summary>
    public LineStyle InsideCellsBorderStyle { get; set; } = LineStyle.Thin;

    /// <summary>
    /// Set color of inside borders of table Cells. It do not effect the table Outside borders! Default is Thin border (Like Excel normal cells)
    /// </summary>
    public KnownColor InsideCellsBorderColor { get; set; } = KnownColor.LightGray;
}