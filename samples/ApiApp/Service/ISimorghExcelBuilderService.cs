using ApiApp.SimorghReportModels;
using ExcelWizard.Models;

namespace ApiApp.Service;

public interface ISimorghExcelBuilderService
{
    /// <summary>
    /// Generate Simorgh Co. VoucherStatement Excel Report with Custom Template from VoucherStatementResult model.
    /// With this Service, We Can Generate The Customized Excel everywhere in the project by just calling this method
    /// </summary>
    /// <returns> Returns ByteArray of Generated Excel file </returns>
    public GeneratedExcelFile GenerateVoucherStatementExcelReport(VoucherStatementResult voucherStatement);

    /// <summary>
    /// Generate Simorgh Co. VoucherStatement Excel Report with Custom Template from VoucherStatementResult model.
    /// With this Service, We Can Generate The Customized Excel everywhere in the project by just calling this method
    /// </summary>
    /// <returns> Returns Full Path which the Generated Excel Saved there </returns>
    public string GenerateVoucherStatementExcelReport(VoucherStatementResult voucherStatement, string savePath);
}