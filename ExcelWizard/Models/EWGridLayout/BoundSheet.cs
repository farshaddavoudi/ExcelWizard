using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelWizard.Models.EWGridLayout;

/// <typeparam name="T">Type of the model supposed to be bound</typeparam>
/// <param name="BoundData">Special Data model with ExcelWizard attributes to customize the generated Excel. Should be List of items</param>
/// <param name="SheetName">Name of the bound Sheet. Is not set, get the Sheet name from SheetName property of [ExcelSheet] attribute</param>
public record BoundSheet<T>(List<T> BoundData, string? SheetName = default);

/// <summary>
/// Required data to create a particular sheet by binding it to a model.
/// It helps in scenarios when we want to have multiple different sheets (both differ in binding model type and sheet name) in generated Excel
/// </summary>
/// <param name="BoundData">List of data. e.g. object list of Persons, Phones, Universities, etc which each will be mapped to a Sheet</param>
/// <param name="SheetName">Name of the bound Sheet. If not set, get the Sheet name from SheetName property of [ExcelSheet] attribute</param>
public record BoundSheet(object BoundData, string? SheetName = default) : IValidatableObject
{
    public void ValidateBoundSheetInstance()
    {
        if (BoundData is not IEnumerable)
            throw new ValidationException("Object provided for Sheet binding should be a collection of records");
    }

    // Validations
    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        if (BoundData is not IEnumerable)
            yield return new ValidationResult("Object provided for Sheet binding should be a collection of records", new List<string> { nameof(BoundData) });
    }
}