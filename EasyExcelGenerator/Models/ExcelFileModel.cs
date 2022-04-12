using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EasyExcelGenerator.Models
{
    public class ExcelFileModel
    {
        /// <summary>
        /// Excel file will be generated with this file name
        /// </summary>
        [Required(ErrorMessage = "FileName is required")]
        public string? FileName { get; set; }

        /// <summary>
        /// Sheets shared default styles including default ColumnWidth, default RowHeight and sheets language Direction
        /// </summary>
        public SheetsDefaultStyles SheetsDefaultStyles { get; set; } = new();

        /// <summary>
        /// Set the default IsLocked value for all Sheets
        /// </summary>
        public bool SheetsDefaultIsLocked { get; set; } = false;

        /// <summary>
        /// Excel Sheets model
        /// </summary>
        public List<Sheet> Sheets { get; set; } = new();
    }
}
