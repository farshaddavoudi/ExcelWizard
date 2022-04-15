using EasyExcelGenerator.Models;
using System.Drawing;

namespace ApiApp;

[EasyExcelGridSheet(sheetName: "MyReport", headerBackgroundColor: KnownColor.LightBlue, headerHeight: 20)]
public class AppExcelReportModel
{
    [EasyExcelGridColumn("شناسه")]
    public int Id { get; set; }

    [EasyExcelGridColumn("نام کامل", columnWidthCalculationType: ColumnWidthCalculationType.ExplicitValue, columnWidth: 400)]
    public string? FullName { get; set; }

    [EasyExcelGridColumn("شماره پرسنلی")]
    public string? PersonnelCode { get; set; }
}