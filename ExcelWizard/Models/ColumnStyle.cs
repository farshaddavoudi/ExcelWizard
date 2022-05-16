namespace ExcelWizard.Models;

public class ColumnStyle
{
    /// <param name="columnLetter"> Column Letter </param>
    public ColumnStyle(string columnLetter)
    {
        ColumnNumber = columnLetter.GetCellColumnNumberByCellColumnLetter();
    }

    /// <param name="columnNumber"> Column Number </param>
    public ColumnStyle(int columnNumber)
    {
        ColumnNumber = columnNumber;
    }

    public int ColumnNumber { get; set; }

    public ColumnWidth? ColumnWidth { get; set; } = null; //If not specified, default would be considered

    public TextAlign ColumnTextAlign { get; set; } = TextAlign.Right; //Default RTL direction

    public bool IsColumnHidden { get; set; } = false;

    // TODO: Add MergedCells for Columns property

    public bool AutoFit { get; set; } = false; //TODO: has same concept with Width class (duplicate)

    public bool? IsColumnLocked { get; set; } = null; //Default is null, and it gets Sheet "IsLocked" property value in this case, but if specified, it will override it
}