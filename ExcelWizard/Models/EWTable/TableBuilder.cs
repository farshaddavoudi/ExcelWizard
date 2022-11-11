using System.Collections.Generic;

namespace ExcelWizard.Models.EWTable;

public static class TableBuilder
{
    private static Table _table = new();

    /// <summary>
    /// Automatically build the Table using the model data and attributes. Model should be an IEnumerable object i.e. list of an item
    /// </summary>
    /// <param name="dataList"></param>
    public static void SetUsingModel(object dataList)
    {
        _table.MergedCellsList = new List<MergedCells>();
    }

    /// <summary>
    /// Manually build the Table defining its properties and styles one by one
    /// </summary>
    public static void SetManually()
    {

    }

    public static Table Build()
    {
        return _table;
    }
}