using EasyExcelGenerator.Models;
using System.Drawing;

namespace ApiApp;

[ExcelSheet(SheetName = "MyUsers", DefaultTextAlign = TextAlign.Center, HeaderBackgroundColor = KnownColor.LightBlue, HeaderHeight = 40,
    BorderType = LineStyle.DashDotDot, DataBackgroundColor = KnownColor.Bisque, DataRowHeight = 25, IsSheetLocked = true,
    SheetDirection = SheetDirection.RightToLeft, FontColor = KnownColor.Red)]
public class User
{
    [ExcelColumn(HeaderName = "UserId", HeaderTextAlign = TextAlign.Right, DataTextAlign = TextAlign.Right, FontColor = KnownColor.Blue)]
    public int Id { get; set; }

    [ExcelColumn(HeaderName = "Name", HeaderTextAlign = TextAlign.Left, FontWeight = FontWeight.Bold)]
    public string? FullName { get; set; }

    [ExcelColumn(HeaderName = "Personnel No", HeaderTextAlign = TextAlign.Left, ColumnWidth = 50, FontSize = 15)]
    public string? PersonnelCode { get; set; }

    public string? Nationality { get; set; }
}