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
        // have proper comments to make them clear

        // Building Excel Part: Seeing the Excel template At a glance, we can see it is created of below parts:
        // 1- Top header which is a Table (is not a Row because of occupying two Rows) which is Merged and became a Unite Cell
        // 2- Having a Row which is the first table Header (It can be part of the Table model, but it makes it a little hard because
        // the Table data is dynamic and it is better to see the Table header as a single Row.
        // 3- First table with some dynamic data which the data is currency type
        // 4- Again having a Row which is the second table Header.
        // 5- Second table again with dynamic data.
        // 6- A Table which actually is the bottom data header. It is merged, so cannot be declared as a Row.
        // 7- Bottom table with think inside border having حقوق پایه و اضافه کار Data in it
        // 8- Last Table for Bottom data which again is merged


        // Table: Excel Header 
        var tableHeader = new Table
        {
            TableRows = new List<Row>
            {
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("A", 1) { Value = "گزارش تست", CellStyle = new CellStyle
                        {
                            CellTextAlign = TextAlign.Center
                        }}
                    }
                }
            },
            TableStyle = new(),
            MergedCellsList = new List<MergedCells>
            {
                new()
                {
                    MergedBoundaryLocation = new()
                    {
                        FirstCellLocation = new CellLocation("A", 1),
                        LastCellLocation = new CellLocation("H", 2)
                    }
                }
            }
        };

        // Gray bg row (کد حساب - بدهکار - بستانکار) - First table Header
        var firstTableHeaderRow = new Row
        {
            RowCells = new List<Cell>
            {
                new("A", 3) {Value = "کد حساب"},
                new("B", 3) {Value = "بدهکار"},
                new("C", 3) {Value = "بستانکار"}
            },

            RowStyle = new RowStyle
            {
                BackgroundColor = Color.Gray
            }
        };

        // First table with header of (کد حساب - بدهکار - بستانکار)
        var firstTable = new Table
        {
            TableRows = voucherStatement.VoucherStatementItem.Select((item, index) => new Row
            {
                RowCells = new List<Cell>
                {
                    new("A", index + 4) {Value = item.AccountCode},
                    new("B", index + 4) {Value = item.Debit, CellContentType = CellContentType.Currency},
                    new("C", index + 4) {Value = item.Credit, CellContentType = CellContentType.Currency}
                }
            }).ToList(),

            TableStyle = new TableStyle
            {
                OutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                CellsSeparatorBorder = new Border { BorderLineStyle = LineStyle.Thick }
            }
        };

        // Gray bg row (کد حساب - بدهکار - بستانکار) - Second table Header
        var secondTableHeaderRow = new Row
        {
            RowCells = new List<Cell>
            {
                new("A", firstTable.GetNextVerticalRowNumberAfterTable()) {Value = "کد حساب"},
                new("B", firstTable.GetNextVerticalRowNumberAfterTable()) {Value = "بدهکار"},
                new("C", firstTable.GetNextVerticalRowNumberAfterTable()) {Value = "بستانکار"}
            },

            RowStyle = new RowStyle
            {
                BackgroundColor = Color.Gray
            }
        };

        // Second table with header of (کد حساب - بدهکار - بستانکار)
        var secondTable = new Table
        {
            TableRows = voucherStatement.VoucherStatementItem.Select((item, index) => new Row
            {
                RowCells = new List<Cell>
                {
                    new("A", index + secondTableHeaderRow.GetNextRowNumberAfterRow()) { Value = item.AccountCode },
                    new("B", index + secondTableHeaderRow.GetNextRowNumberAfterRow()) { Value = item.Debit, CellContentType = CellContentType.Currency },
                    new("C", index + secondTableHeaderRow.GetNextRowNumberAfterRow()) { Value = item.Credit, CellContentType = CellContentType.Currency }
                }
            }).ToList(),

            TableStyle = new TableStyle
            {
                OutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                CellsSeparatorBorder = new Border { BorderLineStyle = LineStyle.Thick }
            }
        };

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
                        tableHeader,
                        firstTable,
                        secondTable
                    },

                    SheetRows = new()
                    {
                        firstTableHeaderRow,
                        secondTableHeaderRow
                    },

                    SheetCells = new()
                }
            }
        };
    }
}