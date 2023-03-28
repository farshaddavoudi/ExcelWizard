using ExcelWizard.Models.EWStyles;
using System;
using System.Drawing;

namespace ExcelWizard.Models.EWGridLayout;

/// <summary>
/// Configure the Excel generic properties in a grid layout Excel schema
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ExcelSheetAttribute : Attribute
{
    /// <summary>
    /// Sheet direction 
    /// </summary>
    public SheetDirection SheetDirection { get; set; } = SheetDirection.LeftToRight;

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
    /// Sheet Cells FontFamily Name
    /// </summary>
    public string? FontName { get; set; }

    /// <summary>
    /// Sheet Cells FontColor
    /// </summary>
    public KnownColor FontColor { get; set; } = KnownColor.Black;

    /// <summary>
    /// Sheets Cells FontSize. If 0 it means default FontSize
    /// </summary>
    public int FontSize { get; set; }

    /// <summary>
    /// Font Weight for entire Sheet. Inherit is equal to default here which is bold for Header and normal for Cells
    /// </summary>
    public FontWeight FontWeight { get; set; } = FontWeight.Inherit;

    /// <summary>
    /// Sheet All Data Cells Background
    /// </summary>
    public KnownColor DataBackgroundColor { get; set; } = KnownColor.Transparent;

    /// <summary>
    /// All Borders Type
    /// </summary>
    public LineStyle BorderType { get; set; } = LineStyle.Thin;

    /// <summary>
    /// All Borders Color
    /// </summary>
    public KnownColor BorderColor { get; set; } = KnownColor.LightGray;

    /// <summary>
    /// Are Sheet Cells Locked? Meaning you cannot edit/delete Cells data but the Sheet can still be formatted
    /// </summary>
    public bool IsSheetLocked { get; set; }

    /// <summary>
    /// The Sheet will be hardly protected and you cannot format/delete Cells/Rows/Columns or edit any objects
    /// </summary>
    public bool IsSheetHardProtected { get; set; }
}