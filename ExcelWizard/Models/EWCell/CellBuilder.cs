using System;

namespace ExcelWizard.Models.EWCell;

public class CellBuilder
{
    private CellBuilder() { }

    private Cell Cell { get; set; } = new();
    private static bool CanBuild { get; set; }

    /// <summary>
    /// Set Cell location (X/Column and Y/Row)
    /// </summary>
    /// <param name="columnLetterOrNumber"> [X]; Cell Column Letter (e.g. "B") or Cell Column Number (e.g. 2) </param>
    /// <param name="rowNumber"> [Y]; Cell Row Number </param>
    public static CellBuilder SetLocation(dynamic columnLetterOrNumber, int rowNumber)
    {
        CanBuild = true;

        return new CellBuilder
        {
            Cell = new Cell
            {
                CellLocation = new CellLocation(columnLetterOrNumber, rowNumber)
            }
        };
    }

    /// <summary>
    /// Set Cell Value that is displayed
    /// </summary>
    /// <param name="value">Cell Content Value</param>
    public CellBuilder SetValue(object? value)
    {
        Cell.CellValue = value;

        return this;
    }

    /// <summary>
    /// Set an arbitrary property to distinguish the Cells. For example can be the db Id (which are not suppose to be shown in the Excel)
    /// </summary>
    /// <param name="identifier">Cell unique identifier</param>
    public CellBuilder SetIdentifier(string identifier)
    {
        Cell.CellIdentifier = identifier;

        return this;
    }

    /// <summary>
    /// Set content type of Cell e.g. text, currency, date, number, etc.
    /// </summary>
    /// <param name="contentType"></param>
    public CellBuilder SetContentType(CellContentType contentType)
    {
        Cell.CellContentType = contentType;

        return this;
    }

    /// <summary>
    /// Set Cell Styles including Font, Wrap behaviour, Align and etc
    /// </summary>
    public CellBuilder SetCellStyle(CellStyle cellStyle)
    {
        Cell.CellStyle = cellStyle;

        return this;
    }

    /// <summary>
    /// Show / Hide Cell Content in Generated Excel
    /// </summary>
    /// <param name="isVisible"></param>
    public CellBuilder SetVisibility(bool isVisible)
    {
        Cell.IsCellVisible = isVisible;

        return this;
    }

    /// <summary>
    /// Will override the IsSheetLocked property of Sheet model if set to a value. Default will inherit
    /// </summary>
    public CellBuilder SetLockStatus(bool isLocked)
    {
        Cell.IsCellLocked = isLocked;

        return this;
    }

    public Cell Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Cell because its location is undefined");

        return Cell;
    }
}