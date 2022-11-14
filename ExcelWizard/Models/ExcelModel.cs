using ExcelWizard.Models.EWSheet;
using System.Collections.Generic;

namespace ExcelWizard.Models;

public class ExcelModel
{
    /// <summary>
    /// Excel file name without .Xlsx extension. Excel file will be generated with this file name
    /// </summary>
    public string? GeneratedFileName { get; internal set; }

    /// <summary>
    /// All Sheets shared default styles including default ColumnWidth, default RowHeight and sheets language Direction
    /// </summary>
    public SheetsDefaultStyle SheetsDefaultStyle { get; internal set; } = new();

    /// <summary>
    /// Set the default IsLocked value for all Sheets. Default is false
    /// </summary>
    public bool AreSheetsLockedByDefault { get; internal set; } = false;

    /// <summary>
    /// Excel Sheets data and configurations
    /// </summary>
    public List<Sheet> Sheets { get; internal set; } = new();
}