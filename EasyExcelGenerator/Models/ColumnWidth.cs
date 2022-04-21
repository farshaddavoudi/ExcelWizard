using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EasyExcelGenerator.Models;

public class ColumnWidth : IValidatableObject
{
    public ColumnWidthCalculationType WidthCalculationType { get; set; } = ColumnWidthCalculationType.AdjustToContents;

    /// <summary>
    /// Width value of the Column. In case of ColumnWidthCalculationType.AdjustToContents it should be left null
    /// </summary>
    public double? Width { get; set; }

    // Validations
    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        if (WidthCalculationType == ColumnWidthCalculationType.AdjustToContents)
        {
            if (Width is not null)
            {
                yield return new ValidationResult(
                    "Column with AdjustToContent Width calculation type cannot have explicit value",
                    new List<string> { nameof(Width) });
            }
        }

        if (WidthCalculationType == ColumnWidthCalculationType.ExplicitValue)
        {
            if (Width is null)
            {
                yield return new ValidationResult(
                    "Column width value should be specified when CalculationType is set to explicit value",
                    new List<string> { nameof(Width) }
                );
            }
        }
    }
}