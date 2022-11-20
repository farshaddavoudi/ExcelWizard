namespace ExcelWizard.Models.EWExcel;

public class BindingSheet
{
    /// <param name="bindingListModel">List of data list. e.g. object list of Persons, Phones, Universities, etc which each will be mapped to a Sheet</param>
    /// <param name="sheetName">Name of the bound Sheet. If not set, get the Sheet name from SheetName property of [ExcelSheet] attribute</param>
    public BindingSheet(object bindingListModel, string? sheetName = null)
    {
        BindingListModel = bindingListModel;
        SheetName = sheetName;
    }

    /// <summary>
    /// List of data list. e.g. object list of Persons, Phones, Universities, etc which each will be mapped to a Sheet
    /// </summary>
    public object? BindingListModel { get; set; }

    /// <summary>
    /// Name of the bound Sheet. If not set, get the Sheet name from SheetName property of [ExcelSheet] attribute
    /// </summary>
    public string? SheetName { get; set; }
}