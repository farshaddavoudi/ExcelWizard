using System;

namespace ExcelWizard.Models;

public class CellLocation
{
    /// <param name="x"> Cell Column Letter </param>
    /// <param name="rowNumber"> Cell Row Number </param>
    public CellLocation(string x, int rowNumber)
    {
        ColumnNumber = GetCellColumnNumberByCellColumnLetter(x);
        RowNumber = rowNumber;
    }

    /// <param name="columnNumber"> Cell Column Number </param>
    /// <param name="rowNumber"> Cell Row Number </param>
    public CellLocation(int columnNumber, int rowNumber)
    {
        ColumnNumber = columnNumber;
        RowNumber = rowNumber;
    }

    public int ColumnNumber { get; set; }

    public int RowNumber { get; set; }

    /// <summary>
    /// Get Cell Location Display Name, e.g. "A2" or "B13"
    /// </summary>
    /// <returns></returns>
    public string GetCellLocationDisplayName()
    {
        return $"{GetCellColumnLetterByCellColumnNumber(ColumnNumber)}{RowNumber}";
    }

    /// <summary>
    ///  Get Cell Column Number from Cell Column Letter, e.g. "A" => 1 or "C" => 3
    /// </summary>
    private int GetCellColumnNumberByCellColumnLetter(string cellColumnLetter)
    {
        int retVal = 0;
        string col = cellColumnLetter.ToUpper();
        for (int iChar = col.Length - 1; iChar >= 0; iChar--)
        {
            char colPiece = col[iChar];
            int colNum = colPiece - 64;
            retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
        }
        return retVal;
    }

    /// <summary>
    /// Get Cell Column Letter By Cell Column Number, e.g. 1 => "A" or 3 => "C"
    /// </summary>
    /// <param name="cellColumnNumber"></param>
    /// <returns></returns>
    private string GetCellColumnLetterByCellColumnNumber(int cellColumnNumber)
    {
        int dividend = cellColumnNumber;

        string cellName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            cellName = Convert.ToChar(65 + modulo) + cellName;
            dividend = (dividend - modulo) / 26;
        }

        return cellName.ToUpper();
    }
}