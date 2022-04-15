using System;

namespace EasyExcelGenerator.Models;

[AttributeUsage(AttributeTargets.Property)]
public class EasyExcelGridColumnAttribute : Attribute
{
    public EasyExcelGridColumnAttribute(string? headerName = null,
        CellType excelDataType = CellType.Text,
        ColumnWidthCalculationType columnWidthCalculationType = ColumnWidthCalculationType.AdjustToContents,
        int columnWidth = 0
        )
    {
        HeaderName = headerName;
        ExcelDataType = excelDataType;
        ColumnWidth = new ColumnWidth
        {
            Width = columnWidthCalculationType == ColumnWidthCalculationType.AdjustToContents ? null : columnWidth,
            WidthCalculationType = columnWidthCalculationType
        };
    }

    public string? HeaderName { get; set; }

    public CellType ExcelDataType { get; set; }

    public ColumnWidth? ColumnWidth { get; set; } //TODO: this property don't work!!!
}