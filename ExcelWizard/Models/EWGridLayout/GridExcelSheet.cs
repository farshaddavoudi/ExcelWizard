using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelWizard.Models.EWGridLayout;

public class GridExcelSheet
{
    /// <summary>
    /// Special Data model with ExcelWizard attributes to customize the generated Excel. Should be List of items
    /// </summary>
    [Required]
    public object? DataList { get; set; }


    // Validations
    public void ValidateGridExcelSheetInstance()
    {
        if (DataList is not IEnumerable)
            throw new ValidationException("Object provided for GridExcelSheet should be a Collection of records");
    }
}