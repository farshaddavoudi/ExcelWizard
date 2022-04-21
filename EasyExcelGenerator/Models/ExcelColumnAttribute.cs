using System;

namespace EasyExcelGenerator.Models;

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
}