using EasyExcelGenerator.Models;

namespace EasyExcelGenerator.Service;

public interface IEasyExcelService
{
    /// <summary>
    /// Generate Simple Grid Excel file from special model configured options with EasyExcel attributes
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"></param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    /// <summary>
    /// Generate Simple Single Sheet Grid Excel file from special model configured options with EasyExcel attributes
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List<Person>()) which you want to show as Excel Report </param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateGridLayoutExcel(object singleSheetDataList);

    /// <summary>
    /// Generate Grid Layout Excel having multiple Sheets from special model configured options with EasyExcel attributes
    /// Save it in path and return the saved url
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"> Model for Multiple Sheets Grid Layout Excel. For Single Sheet, use another overload with object arg </param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public string GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder, string savePath);

    /// <summary>
    /// Generate Simple Single Sheet Grid Excel file from special model configured options with EasyExcel attributes
    /// Save it in path and return the saved url
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List<Person>()) which you want to show as Excel Report </param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public string GenerateGridLayoutExcel(object singleSheetDataList, string savePath);

    /// <summary>
    /// Generate Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location
    /// </summary>
    /// <param name="compoundExcelBuilder"></param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    /// <summary>
    /// Generate Excel file and save it in path and return the saved url
    /// </summary>
    /// <param name="compoundExcelBuilderFile"></param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public string GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilderFile, string savePath);
}