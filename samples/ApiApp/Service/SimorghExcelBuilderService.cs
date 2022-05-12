using ApiApp.SimorghReportModels;
using ExcelWizard.Models;
using ExcelWizard.Service;
using System.Drawing;

namespace ApiApp.Service;

public class SimorghExcelBuilderService : ISimorghExcelBuilderService
{
    private readonly IExcelWizardService _excelWizardService;

    // DI
    public SimorghExcelBuilderService(IExcelWizardService excelWizardService)
    {
        _excelWizardService = excelWizardService;
    }


    public GeneratedExcelFile GenerateVoucherStatementExcelReport(VoucherStatementResult voucherStatement)
    {
        var excelModel = GetExcelModelFromVoucherStatementResult(voucherStatement);

        return _excelWizardService.GenerateCompoundExcel(excelModel);
    }

    public string GenerateVoucherStatementExcelReport(VoucherStatementResult voucherStatement, string savePath)
    {
        var excelModel = GetExcelModelFromVoucherStatementResult(voucherStatement);

        return _excelWizardService.GenerateCompoundExcel(excelModel, savePath);
    }

    /// <summary>
    /// Defined to use for both methods and do not duplicate codes
    /// </summary>
    private CompoundExcelBuilder GetExcelModelFromVoucherStatementResult(VoucherStatementResult voucherStatement)
    {
        // It is the heart of using the ExcelWizard package to generate your desired Excel report
        // You should create your Excel template (CompoundExcelBuilder model) using your local app model (here VoucherStatementResult)
        // Just start with CompoundExcelBuilder and the properties names speak for themselves. Also note all properties
        // -have proper comments to make them clear

        return new CompoundExcelBuilder
        {
            GeneratedFileName = voucherStatement.ReportName,

            Sheets = new List<Sheet>
            {
                new()
                {
                    SheetName = "RemainReport",

                    SheetTables = new(),

                    SheetRows = new()
                    {
                        // Gray bg row (کد حساب - بدهکار - بستانکار) - row no 3
                        new Row
                        {
                            RowStyle = new RowStyle
                            {
                                BackgroundColor = Color.Gray
                            }
                        }
                    },

                    SheetCells = new()
                }
            }
        };
    }
}