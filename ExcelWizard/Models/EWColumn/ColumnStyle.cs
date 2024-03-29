﻿using ExcelWizard.Models.EWStyles;

namespace ExcelWizard.Models.EWColumn;

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

    /// <summary>
    /// If not specified, default would be considered
    /// </summary>
    public ColumnWidth? ColumnWidth { get; set; } = null;

    public TextAlign ColumnTextAlign { get; set; } = TextAlign.Right; //Default RTL direction

    public bool IsColumnHidden { get; set; } = false;

    // TODO: Add MergedCells for Columns property

    public bool AutoFit { get; set; } = false; //TODO: has same concept with Width class (duplicate)

    /// <summary>
    /// Default is null, and it gets Sheet "IsLocked" property value in this case, but if specified, it will override it
    /// </summary>
    public bool? IsColumnLocked { get; set; } = null;
}