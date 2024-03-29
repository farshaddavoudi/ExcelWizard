﻿namespace ExcelWizard.Models.EWCell;

public class CellLocation
{
    /// <param name="columnLetter"> Cell Column Letter </param>
    /// <param name="rowNumber"> Cell Row Number </param>
    public CellLocation(string columnLetter, int rowNumber)
    {
        ColumnNumber = columnLetter.GetCellColumnNumberByCellColumnLetter();
        RowNumber = rowNumber;
    }

    /// <param name="columnNumber"> Cell Column Number </param>
    /// <param name="rowNumber"> Cell Row Number </param>
    public CellLocation(int columnNumber, int rowNumber)
    {
        ColumnNumber = columnNumber;
        RowNumber = rowNumber;
    }

    /// <summary>
    /// X Location
    /// </summary>
    public int ColumnNumber { get; }

    /// <summary>
    /// Y Location
    /// </summary>
    public int RowNumber { get; }

    /// <summary>
    /// Get Cell Location Display Name, e.g. "A2" or "B13"
    /// </summary>
    /// <returns></returns>
    public string GetCellLocationDisplayName()
    {
        return $"{ColumnNumber.GetCellColumnLetterByCellColumnNumber()}{RowNumber}";
    }
}