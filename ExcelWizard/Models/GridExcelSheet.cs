using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelWizard.Models;

public class GridExcelSheet : IValidatableObject
{
    /// <summary>
    /// Special Data model with ExcelWizard attributes to customize the generated Excel. Should be List of items
    /// </summary>
    [Required]
    public object? DataList { get; set; }


    // Validations
    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        if (DataList is not IEnumerable)
            yield return new ValidationResult("Object provided for GridExcelSheet should be a Collection of records");
    }
}