namespace ExcelWizard.Models;

public class Cell
{
    #region Ctors

    /// <param name="columnLetter"> Cell Column Letter (X) </param>
    /// <param name="rowNumber"> Cell Row Number (Y) </param>
    public Cell(string columnLetter, int rowNumber)
    {
        CellLocation = new CellLocation(columnLetter, rowNumber);
    }

    /// <param name="columnNumber"> Cell Column Number (X) </param>
    /// <param name="rowNumber"> Cell Row Number (Y) </param>
    public Cell(int columnNumber, int rowNumber)
    {
        CellLocation = new CellLocation(columnNumber, rowNumber);
    }

    /// <param name="columnLetter"> Cell Column Letter (X) </param>
    /// <param name="rowNumber"> Cell Row Number (Y) </param>
    /// <param name="cellValue"> Cell Content Value </param>
    public Cell(string columnLetter, int rowNumber, object? cellValue)
    {
        CellLocation = new CellLocation(columnLetter, rowNumber);

        CellValue = cellValue;
    }

    /// <param name="columnNumber"> Cell Column Number (X) </param>
    /// <param name="rowNumber"> Cell Row Number (Y) </param>
    /// <param name="cellValue"> Cell Content Value </param>
    public Cell(int columnNumber, int rowNumber, object? cellValue)
    {
        CellLocation = new CellLocation(columnNumber, rowNumber);

        CellValue = cellValue;
    }

    #endregion

    /// <summary>
    /// An arbitrary property to distinguish the Cells. For example can be the db Id (which are not suppose to be shown in the Excel)
    /// </summary>
    public string? CellIdentifier { get; set; }

    /// <summary>
    /// Cell Value that are displayed
    /// </summary>
    public object? CellValue { get; set; }

    /// <summary>
    /// Cell Location. Row is Number only and Column can be both Letter (e.g. "B") or No (e.g. 2)
    /// </summary>
    public CellLocation CellLocation { get; set; }

    /// <summary>
    /// Set Cell Styles including Font, Wrap behaviour, Align and etc
    /// </summary>
    public CellStyle CellStyle { get; set; } = new();

    public CellContentType CellContentType { get; set; } = CellContentType.General;

    /// <summary>
    /// Show / Hide Cell Content in Generated Excel
    /// </summary>
    public bool IsCellVisible { get; set; } = true;

    /// <summary>
    /// Will override the IsSheetLocked property of Sheet model if set to a value. Default will inherit
    /// </summary>
    public bool? IsCellLocked { get; set; } = null;
}