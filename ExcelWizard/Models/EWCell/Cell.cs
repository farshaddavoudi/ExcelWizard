namespace ExcelWizard.Models.EWCell;

public class Cell
{
    /// <summary>
    /// An arbitrary property to distinguish the Cells. For example can be the db Id (which are not suppose to be shown in the Excel)
    /// </summary>
    public string? CellIdentifier { get; internal set; }

    /// <summary>
    /// Cell Value that are displayed
    /// </summary>
    public object? CellValue { get; internal set; }

    /// <summary>
    /// Cell Location. Row is Number only and Column can be both Letter (e.g. "B") or No (e.g. 2)
    /// </summary>
#pragma warning disable CS8618
    public CellLocation CellLocation { get; internal set; }
#pragma warning restore CS8618

    /// <summary>
    /// Set Cell Styles including Font, Wrap behaviour, Align and etc
    /// </summary>
    public CellStyle CellStyle { get; internal set; } = new();

    public CellContentType CellContentType { get; internal set; } = CellContentType.General;

    /// <summary>
    /// Show / Hide Cell Content in Generated Excel
    /// </summary>
    public bool IsCellVisible { get; internal set; } = true;

    /// <summary>
    /// Will override the IsSheetLocked property of Sheet model if set to a value. Default will inherit
    /// </summary>
    public bool? IsCellLocked { get; internal set; } = null;
}