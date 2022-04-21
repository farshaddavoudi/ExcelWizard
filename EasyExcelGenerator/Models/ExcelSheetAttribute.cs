using System;
using System.Drawing;

namespace EasyExcelGenerator.Models;

/// <summary>
/// Configure the Excel generic properties
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ExcelSheetAttribute : Attribute
{
    /// <summary>
    /// Sheet direction 
    /// </summary>
    public SheetDirection SheetDirection { get; set; } = SheetDirection.RightToLeft;

    /// <summary>
    /// Sheet name of generated Excel that contains the Class data. The default is Sheet1, Sheet2, etc..
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// Default text align including both header and data cells. It can be overridden for header as well as data cells
    /// </summary>
    public TextAlign DefaultTextAlign { get; set; } = TextAlign.Center;

    /// <summary>
    /// Sheet Header Height. 0 will revert to default
    /// </summary>
    public int HeaderHeight { get; set; }

    /// <summary>
    /// Sheet Header Background Color
    /// </summary>
    public KnownColor HeaderBackgroundColor { get; set; } = KnownColor.Transparent;

    /// <summary>
    /// Sheet Each Data Row Height
    /// </summary>
    public int DataRowHeight { get; set; }

    /// <summary>
    /// Sheet All Data Cells Background
    /// </summary>
    public KnownColor DataBackgroundColor { get; set; } = KnownColor.Transparent;

    /// <summary>
    /// All Borders Type
    /// </summary>
    public LineStyle BorderType { get; set; } = LineStyle.Thin;

    /// <summary>
    /// Are Sheet Cells Locked? Meaning you cannot edit/delete Cells data but the Sheet can still be formatted
    /// </summary>
    public bool IsSheetLocked { get; set; }

    /// <summary>
    /// The Sheet will be hardly protected and you cannot format/delete Cells/Rows/Columns or edit any objects
    /// </summary>
    public bool IsSheetHardProtected { get; set; }
}