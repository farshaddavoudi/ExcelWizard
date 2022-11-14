using BlazorDownloadFile;
using ExcelWizard.Models;
using System.Threading.Tasks;

namespace ExcelWizard.Service;

public interface IExcelWizardService
{
    /// <summary>
    /// Generate Excel file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelModel"> ExcelModel created by ExcelBuilder </param>
    /// <returns> Byte array of generated Excel saved in memory. </returns>
    GeneratedExcelFile GenerateExcel(ExcelModel excelModel);

    /// <summary>
    /// Generate Excel file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelModel"> ExcelModel created by ExcelBuilder </param>
    /// <param name="savePath"> The url saved </param>
    /// <returns> Save generated Excel in a path in your device </returns>
    string GenerateExcel(ExcelModel excelModel, string savePath);

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the generated file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelModel">  ExcelModel created by ExcelBuilder </param>
    /// <returns> Instantly download from Browser </returns>
    Task<DownloadFileResult> GenerateAndDownloadExcelInBlazor(ExcelModel excelModel);
}