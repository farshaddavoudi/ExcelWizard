using EasyExcelGenerator.Models;
using System.Drawing;

namespace ApiApp;

[ExcelSheet(SheetName = "MyReport", DefaultTextAlign = TextAlign.Center, HeaderBackgroundColor = KnownColor.LightBlue, HeaderHeight = 40,
    BorderType = LineStyle.DashDotDot, DataBackgroundColor = KnownColor.Bisque, DataRowHeight = 25, IsSheetLocked = true)]
public class AppExcelReportModel
{
    [ExcelColumn(HeaderName = "شناسه", HeaderTextAlign = TextAlign.Right, DataTextAlign = TextAlign.Right)]
    public int Id { get; set; }

    [ExcelColumn(HeaderName = "Name", HeaderTextAlign = TextAlign.Left)]
    public string? FullName { get; set; }

    [ExcelColumn(HeaderName = "شماره پرسنلی", HeaderTextAlign = TextAlign.Left, ColumnWidth = 50)]
    public string? PersonnelCode { get; set; }
}