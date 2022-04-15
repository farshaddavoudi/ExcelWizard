using System;
using System.Drawing;

namespace EasyExcelGenerator.Models;

[AttributeUsage(AttributeTargets.Class)]
public class EasyExcelGridSheetAttribute : Attribute
{
    #region Constructor

    public EasyExcelGridSheetAttribute(string? sheetName = null, int headerHeight = 0, KnownColor headerBackgroundColor = KnownColor.Transparent)
    {
        SheetName = sheetName;
        HeaderHeight = headerHeight == 0 ? null : headerHeight;
        HeaderBackgroundColor = Color.FromKnownColor(headerBackgroundColor);
    }

    #endregion

    public string? SheetName { get; set; }

    public double? HeaderHeight { get; set; }

    public Color HeaderBackgroundColor { get; set; }
}