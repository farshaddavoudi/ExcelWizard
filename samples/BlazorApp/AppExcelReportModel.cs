using ExcelWizard.Models;
using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace BlazorApp;

[ExcelSheet(SheetName = "MyReport", DefaultTextAlign = TextAlign.Center, HeaderBackgroundColor = KnownColor.LightBlue, HeaderHeight = 40,
    BorderType = LineStyle.DashDotDot, DataBackgroundColor = KnownColor.Bisque, DataRowHeight = 25, IsSheetLocked = true,
    SheetDirection = SheetDirection.RightToLeft, FontColor = KnownColor.Red)]
public class AppExcelReportModel
{
    [ExcelSheetColumn(HeaderName = "شناسه", HeaderTextAlign = TextAlign.Right, DataTextAlign = TextAlign.Right, FontColor = KnownColor.Blue)]
    public int Id { get; set; }

    [ExcelSheetColumn(Ignore = true, HeaderName = "Name", HeaderTextAlign = TextAlign.Left, FontWeight = FontWeight.Bold)]
    public string? FullName { get; set; }

    [ExcelSheetColumn(HeaderName = "شماره پرسنلی", HeaderTextAlign = TextAlign.Left, ColumnWidth = 50, FontSize = 15)]
    public string? PersonnelCode { get; set; }
}