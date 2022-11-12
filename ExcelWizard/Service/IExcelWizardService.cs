using BlazorDownloadFile;
using ExcelWizard.Models;
using ExcelWizard.Models.EWGridLayout;
using System.Threading.Tasks;

namespace ExcelWizard.Service;

public interface IExcelWizardService
{
    /// <summary>
    /// Generate Simple Grid Excel file from special model configured options with ExcelWizard attributes
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"></param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    /// <summary>
    /// Generate Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List{Person}()) </param>
    /// <param name="generatedFileName"> Generated file name. If leave empty, automatically will have a name like ExcelWizardGeneratedFile_2022-04-22 14-06-29 </param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    /// <summary>
    /// Generate Grid Layout Excel having multiple Sheets from special model configured options with ExcelWizard attributes
    /// Save it in path and return the saved url
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"> Model for Multiple Sheets Grid Layout Excel. For Single Sheet, use another overload with object arg </param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public string GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder, string savePath);

    /// <summary>
    /// Generate Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes
    /// Save it in path and return the saved url
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List{Person}()) </param>
    /// <param name="savePath"></param>
    /// <param name="generatedFileName"> Generated file name. If leave empty string, automatically will have a name like ExcelWizardGeneratedFile_2022-04-22 14-06-29 </param>
    /// <returns></returns>
    public string GenerateGridLayoutExcel(object singleSheetDataList, string savePath, string generatedFileName);

    /// <summary>
    /// Generate Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location
    /// </summary>
    /// <param name="compoundExcelBuilder"></param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    /// <summary>
    /// Generate Excel file and save it in path and return the saved url
    /// </summary>
    /// <param name="compoundExcelBuilder"></param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public string GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder, string savePath);


    #region Blazor Application

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the Simple Multiple Sheet Grid Excel file from special model configured options with ExcelWizard attributes in Blazor apps
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"></param>
    /// <returns></returns>
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes in Blazor apps
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List{Person}()) </param>
    /// <param name="generatedFileName"> Generated file name. If leave empty, automatically will have a name like ExcelWizardGeneratedFile_2022-04-22 14-06-29 </param>
    /// <returns></returns>
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location in Blazor apps
    /// </summary>
    /// <param name="compoundExcelBuilder"></param>
    /// <returns></returns>
    public Task<DownloadFileResult> BlazorDownloadCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);



    #endregion
}