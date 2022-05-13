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

        ////////////////////////////////////////
        // 1- Define all Tables in Excel
        ////////////////////////////////////////

        // First table with header of (کد حساب - بدهکار - بستانکار)
        var firstTable = new Table();

        //////////////////////////////////////
        // 2- Define all Rows in Excel
        //////////////////////////////////////

        //////////////////////////////////////
        // 3- Define all Cells in Excel
        //////////////////////////////////////

        return new CompoundExcelBuilder
        {
            GeneratedFileName = voucherStatement.ReportName,

            Sheets = new List<Sheet>
            {
                new()
                {
                    SheetName = "RemainReport",

                    SheetTables = new()
                    {
                        new Table
                        {
                            TableRows = voucherStatement.VoucherStatementItem.Select((item, index) => new Row
                            {
                                RowCells = new List<Cell>
                                {
                                    new("A", index + 4){ Value = item.AccountCode },
                                    new("B", index + 4){ Value = item.Debit, CellContentType = CellContentType.Currency },
                                    new("C", index + 4){ Value = item.Credit, CellContentType = CellContentType.Currency }
                                }
                            }).ToList(),

                            TableStyle = new TableStyle
                            {
                                OutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                                CellsSeparatorBorder = new Border { BorderLineStyle = LineStyle.Thick }
                            }
                        }
                    },

                    SheetRows = new()
                    {
                        // Gray bg row (کد حساب - بدهکار - بستانکار) - row no 3
                        new Row
                        {
                            RowCells = new List<Cell>
                            {
                                new("A", 3) { Value = "کد حساب" },
                                new("B", 3) { Value = "بدهکار" },
                                new("C", 3) { Value = "بستانکار" }
                            },

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