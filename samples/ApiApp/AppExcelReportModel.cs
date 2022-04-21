using EasyExcelGenerator.Models;
using System.Drawing;

namespace ApiApp;

[ExcelSheet("MyReport", headerBackgroundColor: KnownColor.LightBlue, headerHeight: 40, borderType: LineStyle.DashDot, dataBackgroundColor: KnownColor.Bisque, dataRowHeight: 30)]
public class AppExcelReportModel
{
    [ExcelColumn("شناسه", TextAlign.Right, TextAlign.Right)]
    public int Id { get; set; }

    [ExcelColumn("Name", TextAlign.Left, columnWidthCalculationType: ColumnWidthCalculationType.ExplicitValue, columnWidth: 400)]
    public string? FullName { get; set; }

    [ExcelColumn("شماره پرسنلی", dataTextAlign: TextAlign.Left)]
    public string? PersonnelCode { get; set; }
}