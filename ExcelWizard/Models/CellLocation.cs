using System;

namespace ExcelWizard.Models;

public class CellLocation
{
    public CellLocation(string x, int y)
    {
        X = NumberFromExcelCell(x);
        Y = y;
    }

    public CellLocation(int x, int y)
    {
        X = x;
        Y = y;
    }

    public int X { get; set; }

    public int Y { get; set; }

    private int NumberFromExcelCell(string cell)
    {
        int retVal = 0;
        string col = cell.ToUpper();
        for (int iChar = col.Length - 1; iChar >= 0; iChar--)
        {
            char colPiece = col[iChar];
            int colNum = colPiece - 64;
            retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
        }
        return retVal;
    }

    public string GetName()
    {
        return $"{GetExcelCellName(X)}{Y}";
    }

    private string GetExcelCellName(int cellNumber)
    {
        int dividend = cellNumber;

        string cellName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            cellName = Convert.ToChar(65 + modulo) + cellName;
            dividend = (dividend - modulo) / 26;
        }

        return cellName;
    }
}