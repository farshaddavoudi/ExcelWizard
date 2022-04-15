using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EasyExcelGenerator.Models;

public class EasyExcelModel
{
    /// <summary>
    /// Excel file will be generated with this file name
    /// </summary>
    [Required(ErrorMessage = "FileName is required")]
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