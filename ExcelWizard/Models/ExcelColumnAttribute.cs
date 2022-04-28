using System;
using System.Drawing;

namespace ExcelWizard.Models;

/// <summary>
/// Configure the Excel Column mapped to this property
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    /// <summary>
    /// Column Header Name 
    /// </summary>
    public string? HeaderName { get; set; }

    /// <summary>
    /// Header Text Align. Will override default one
    /// </summary>
    public TextAlign HeaderTextAlign { get; set; } = TextAlign.Inherit;

    /// <summary>
    /// Data Cells Text Align for the Column. Will override the default one
    /// </summary>
    public TextAlign DataTextAlign { get; set; } = TextAlign.Inherit;

    /// <summary>
    ///  Excel Data Type. Default is Text type
    /// </summary>
    public CellType ExcelDataType { get; set; } = CellType.Text;

    /// <summary>
    /// Column Width. If 0 it means Width automatically set to AdjustToContents
    /// </summary>
    public int ColumnWidth { get; set; }

    /// <summary>
    /// Column FontFamily Name
    /// </summary>
    public string? FontName { get; set; }

    /// <summary>
    /// Column FontColor. Transparent color means reverting back to Sheet FontColor
    /// </summary>
    public KnownColor FontColor { get; set; } = KnownColor.Transparent;

    /// <summary>
    /// Column FontSize. If 0 it means default FontSize
    /// </summary>
    public int FontSize { get; set; }

    /// <summary>
    /// Is Column Font Bold. Inherit means revert back to Sheet Font Weight (IsBold property)
    /// </summary>
    public FontWeight FontWeight { get; set; } = FontWeight.Inherit;
}