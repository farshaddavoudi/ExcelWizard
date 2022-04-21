using System;

namespace EasyExcelGenerator.Models;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    #region Constructor

    /// <summary>
    /// 
    /// </summary>
    /// <param name="headerName"></param>
    /// <param name="headerTextAlign"></param>
    /// <param name="dataTextAlign"></param>
    /// <param name="excelDataType"></param>
    /// <param name="columnWidthCalculationType"></param>
    /// <param name="columnWidth"></param>
    public ExcelColumnAttribute
    (
        string? headerName = null,
        TextAlign headerTextAlign = TextAlign.Inherit,
        TextAlign dataTextAlign = TextAlign.Inherit,
        CellType excelDataType = CellType.Text,
        ColumnWidthCalculationType columnWidthCalculationType = ColumnWidthCalculationType.AdjustToContents,
        int columnWidth = 0
    )
    {
        HeaderName = headerName;
        HeaderTextAlign = headerTextAlign;
        DataTextAlign = dataTextAlign;
        ExcelDataType = excelDataType;
        ColumnWidth = new ColumnWidth
        {
            Width = columnWidthCalculationType == ColumnWidthCalculationType.AdjustToContents ? null : columnWidth,
            WidthCalculationType = columnWidthCalculationType
        };
    }

    #endregion

    public string? HeaderName { get; set; }

    public TextAlign? HeaderTextAlign { get; set; }

    public TextAlign? DataTextAlign { get; set; }

    public CellType ExcelDataType { get; set; }

    public ColumnWidth? ColumnWidth { get; set; } //TODO: this property don't work!!!
}