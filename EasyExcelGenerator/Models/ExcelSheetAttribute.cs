using System;
using System.Drawing;

namespace EasyExcelGenerator.Models;

[AttributeUsage(AttributeTargets.Class)]
public class ExcelSheetAttribute : Attribute
{
    #region Constructor

    /// <summary>
    /// Configure the Excel generic properties
    /// </summary>
    /// <param name="sheetName"> Sheet name of generated Excel that contains the Class data. The default is Sheet1, Sheet2, etc.. </param>
    /// <param name="defaultTextTextAlign"> Default text align including both header and data cells. It can be overridden for header as well as data cells </param>
    /// <param name="headerHeight"> Sheet Header Height </param>
    /// <param name="headerBackgroundColor"> Sheet Header Background Color </param>
    /// <param name="borderType"> All Borders Type </param>
    /// <param name="dataRowHeight"> Sheet Each Data Row Height </param>
    /// <param name="dataBackgroundColor"> Sheet All Data Cells Background </param>
    public ExcelSheetAttribute(string? sheetName = null, TextAlign defaultTextTextAlign = TextAlign.Center, int headerHeight = 0, KnownColor headerBackgroundColor = KnownColor.Transparent,
        LineStyle borderType = LineStyle.Thin, int dataRowHeight = 0, KnownColor dataBackgroundColor = KnownColor.Transparent)
    {
        SheetName = sheetName;
        DefaultTextAlign = defaultTextTextAlign;
        HeaderHeight = headerHeight == 0 ? null : headerHeight;
        HeaderBackgroundColor = Color.FromKnownColor(headerBackgroundColor);
        BorderType = borderType;
        DataRowHeight = dataRowHeight == 0 ? null : dataRowHeight;
        DataBackgroundColor = Color.FromKnownColor(dataBackgroundColor);
    }

    #endregion

    public string? SheetName { get; set; }

    public TextAlign DefaultTextAlign { get; set; }

    public double? HeaderHeight { get; set; }

    public Color HeaderBackgroundColor { get; set; }

    public double? DataRowHeight { get; set; }

    public Color DataBackgroundColor { get; set; }

    public LineStyle BorderType { get; set; }
}