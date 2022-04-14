using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EasyExcelGenerator.Models;

public class ColumnStyle : IValidatableObject
{
    [Required(ErrorMessage = "ColumnNo is required")]
    public int ColumnNo { get; set; }

    public ColumnWidth? ColumnWidth { get; set; } = null; //If not specified, default would be considered

    public TextAlign TextAlign { get; set; } = TextAlign.Right; //Default RTL direction

    public bool IsHidden { get; set; } = false;

    // TODO: Add MergedCells for Columns property

    public bool AutoFit { get; set; } = false; //TODO: has same concept with Width class (duplicate)

    public bool? IsLocked { get; set; } = null; //Default is null, and it gets Sheet "IsLocked" property value in this case, but if specified, it will override it

    // Validations
    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        if (ColumnNo == default)
            yield return new ValidationResult("ColumnNo is required", new List<string> { nameof(ColumnNo) });
    }
}