using System.Collections.Generic;

namespace EasyExcelGenerator.Models;

public class EasyGridExcelBuilder
{
    public EasyGridExcelBuilder()
    {
        Sheets = new();
    }

    /// <summary>
    /// For faster and easier use in case of single Sheet Excel to be generated
    /// </summary>
    /// <param name="singleSheetDataList"></param>
    public EasyGridExcelBuilder(object singleSheetDataList)
    {
        Sheets = new List<GridExcelSheet>
        {
            new()
            {
                DataList = singleSheetDataList
            }
        };
    }

    public List<GridExcelSheet> Sheets { get; set; }
}