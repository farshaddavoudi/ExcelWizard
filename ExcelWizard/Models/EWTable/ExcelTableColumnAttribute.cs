using ExcelWizard.Models.EWStyles;
using System;
using System.Drawing;

namespace ExcelWizard.Models.EWTable;

/// <summary>
/// Configure the Table Columns 
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelTableColumnAttribute : Attribute
{
    /// <summary>
    /// Ignore the Column from being shown in exported Excel. It also will not occupy any location
    /// </summary>
    public bool Ignore { get; set; }

    /// <summary>
    /// Column Header Name. Default is the property name.
    /// It will be ignored if Table HasHeader property is set to false 
    /// </summary>
    public string? HeaderName { get; set; }

    /// <summary>
    /// Header Text Align. Will override default oneWill be ignored if Table HasHeader property is set to false 
    /// </summary>
    public TextAlign HeaderTextAlign { get; set; } = TextAlign.Inherit;

    /// <summary>
    /// Data Cells Text Align for the Column. Will override the default one
    /// </summary>
    public TextAlign DataTextAlign { get; set; } = TextAlign.Inherit;

    /// <summary>
    ///  Table column data type. Default is Text type
    /// </summary>
    public CellContentType DataContentType { get; set; } = CellContentType.Text;

    /// <summary>
    /// Column FontFamily Name
    /// </summary>
    public string? FontName { get; set; }

    /// <summary>
    /// Column FontColor. Transparent color means reverting back to Table FontColor
    /// </summary>
    public KnownColor FontColor { get; set; } = KnownColor.Transparent;

    /// <summary>
    /// Column FontSize. If 0 it means default FontSize
    /// </summary>
    public int FontSize { get; set; }

    /// <summary>
    /// Is Column Font Bold. Inherit means revert back to Table Font Weight
    /// </summary>
    public FontWeight FontWeight { get; set; } = FontWeight.Inherit;
}