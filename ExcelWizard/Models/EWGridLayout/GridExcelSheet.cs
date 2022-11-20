using System.Collections;
using System.ComponentModel.DataAnnotations;

namespace ExcelWizard.Models.EWGridLayout;

public class GridExcelSheet
{
    /// <summary>
    /// Name of the bound Sheet. Is not set, get the Sheet name from SheetName property of [ExcelSheet] attribute
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// Special Data model with ExcelWizard attributes to customize the generated Excel. Should be List of items
    /// </summary>
    [Required]
    public object? DataList { get; set; }


    // Validations
    public void ValidateGridExcelSheetInstance()
    {
        if (DataList is not IEnumerable)
            throw new ValidationException("Object provided for Sheet binding should be a collection of records");
    }
}