using BlazorDownloadFile;
using ExcelWizard.Models;
using System.Threading.Tasks;

namespace ExcelWizard.Service;

public interface IExcelWizardService
{
    /// <summary>
    /// Generate Excel file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelBuilder"> ExcelBuilder with Build() method at the end </param>
    /// <returns> Byte array of generated Excel saved in memory. </returns>
    GeneratedExcelFile GenerateExcel(IExcelBuilder excelBuilder);

    /// <summary>
    /// Generate Excel file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelBuilder"> ExcelBuilder with Build() method at the end </param>
    /// <param name="savePath"> The url saved </param>
    /// <returns> Save generated Excel in a path in your device </returns>
    string GenerateExcel(IExcelBuilder excelBuilder, string savePath);

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the generated file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelBuilder">  ExcelBuilder with Build() method at the end </param>
    /// <returns> Instantly download from Browser </returns>
    Task<DownloadFileResult> GenerateAndDownloadExcelInBlazor(IExcelBuilder excelBuilder);
}