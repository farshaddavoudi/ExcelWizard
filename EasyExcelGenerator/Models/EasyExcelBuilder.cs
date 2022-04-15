using System.Collections.Generic;

namespace EasyExcelGenerator.Models;

public class EasyExcelBuilder
{
    /// <summary>
    /// Excel file name without .Xlsx extension. Excel file will be generated with this file name
    /// </summary>
    public string? FileName { get; set; }

    /// <summary>
    /// Sheets shared default styles including default ColumnWidth, default RowHeight and sheets language Direction
    /// </summary>
    public AllSheetsDefaultStyle AllSheetsDefaultStyle { get; set; } = new();

    /// <summary>
    /// Set the default IsLocked value for all Sheets
    /// </summary>
    public bool AreSheetsLockedByDefault { get; set; } = false;

    /// <summary>
    /// Excel Sheets data and configurations
    /// </summary>
    public List<Sheet> Sheets { get; set; } = new();
}