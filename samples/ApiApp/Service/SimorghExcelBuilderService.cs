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
        // 6- A Table which actually is the bottom data header (with blue bg). It is merged, so cannot be declared as a Row.
        // 7- Bottom table with thick inside border having حقوق پایه و اضافه کار Data in it
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
        var rowFirstTableHeader = new Row
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
        var table1St = new Table
        {
            TableRows = voucherStatement.VoucherStatementItem.Select((item, index) => new Row
            {
                RowCells = new List<Cell>
                {
                    new("A", index + 4) { Value = item.AccountCode },
                    new("B", index + 4) { Value = item.Debit, CellContentType = CellContentType.Currency },
                    new("C", index + 4) { Value = item.Credit, CellContentType = CellContentType.Currency }
                }
            }).ToList(),

            TableStyle = new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick }
            }
        };

        // Gray bg row (کد حساب - بدهکار - بستانکار) - Second table Header
        var rowSecondTableHeader = new Row
        {
            RowCells = new List<Cell>
            {
                new("A", table1St.GetNextVerticalRowNumberAfterTable()) {Value = "کد حساب"},
                new("B", table1St.GetNextVerticalRowNumberAfterTable()) {Value = "بدهکار"},
                new("C", table1St.GetNextVerticalRowNumberAfterTable()) {Value = "بستانکار"}
            },

            RowStyle = new RowStyle
            {
                BackgroundColor = Color.Gray
            }
        };

        // Second table with header of (کد حساب - بدهکار - بستانکار)
        var table2Nd = new Table
        {
            TableRows = voucherStatement.VoucherStatementItem.Select((item, index) => new Row
            {
                RowCells = new List<Cell>
                {
                    new("A", index + rowSecondTableHeader.GetNextRowNumberAfterRow()) { Value = item.AccountCode },
                    new("B", index + rowSecondTableHeader.GetNextRowNumberAfterRow()) { Value = item.Debit, CellContentType = CellContentType.Currency },
                    new("C", index + rowSecondTableHeader.GetNextRowNumberAfterRow()) { Value = item.Credit, CellContentType = CellContentType.Currency }
                }
            }).ToList(),

            TableStyle = new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick }
            }
        };

        // Bottom section Data Header (Blue Bg)
        // MERGES:
        // The Merges here is very tricky!
        // At first, there are Two Vertical Merges (A17:A18) and (B17:B18). 
        // Then there is one Horizontal Merge (C17:E17) for کارخانه دان
        // The same pattern repeats for پرورش پولت and تخم گزاری جوجه
        // And a Vertical Merge for showing sum (K17:K18) 
        // And a Vertical Merge for Average (L17:L18)
        var tableBottomBlueHeader = new Table
        {
            TableRows = new List<Row>
            {
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("A", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "نام حساب" },
                        new("B", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "کد حساب" },
                        new("C", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "کارخانه دان-51011" },
                        new("D", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "" },
                        new("E", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "" },
                        new("F", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "پرورش پولت-51018" },
                        new("G", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "" },
                        new("H", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "" },
                        new("I", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "تخم گزار تجاری-51035" },
                        new("J", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "" },
                        new("K", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "" },
                        new("L", table2Nd.GetNextVerticalRowNumberAfterTable()) { Value = "میانگین", CellStyle =
                        {
                            BackgroundColor = Color.White,
                            Font = new TextFont { FontColor = Color.Black }
                        }}
                    },
                    RowStyle = new RowStyle { RowHeight = 20 }
                },
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("A", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "" },
                        new("B", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "" },
                        new("C", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "قبل از تسهیم" },
                        new("D", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "بعد از تسهیم" },
                        new("E", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "جمع" },
                        new("F", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "قبل از تسهیم" },
                        new("G", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "بعد از تسهیم" },
                        new("H", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "جمع" },
                        new("I", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "قبل از تسهیم" },
                        new("J", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "بعد از تسهیم" },
                        new("K", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "جمع" },
                        new("L", table2Nd.GetNextVerticalRowNumberAfterTable() + 1) { Value = "", CellStyle =
                        {
                            BackgroundColor = Color.White
                        }}
                    },
                    RowStyle = new RowStyle { RowHeight = 20 }
                }
            },

            TableStyle = new TableStyle
            {
                BackgroundColor = Color.Blue,
                Font = new TextFont { FontColor = Color.White }
            },

            MergedCellsList = new List<MergedCells>
            {
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("A", table2Nd.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("A", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("B", table2Nd.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("B", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("C", table2Nd.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("E", table2Nd.GetNextVerticalRowNumberAfterTable())
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("F", table2Nd.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("H", table2Nd.GetNextVerticalRowNumberAfterTable())
                    }
                }
            }
        };

        // Table with Salaries data with thick borders
        var tableSalaries = new Table
        {
            TableRows = voucherStatement.Accounts.Select((account, index) => new Row
            {
                RowCells = new List<Cell>
                {
                    new ("A", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + index) { Value = account.Name },
                    new ("B", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + index) { Value = account.Code }
                }
            }).ToList(),
            TableStyle = new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black }
            }
        };

        // Last table with sharing before/after data
        var tableSharingBeforeAfterData = new Table
        {
            TableRows = new List<Row>
            {
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("C", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("D", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("E", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("F", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("G", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("H", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("I", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("J", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("K", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                        new("L", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()) { Value = 504000 },
                    }
                }
            },
            TableStyle = new TableStyle { TableTextAlign = TextAlign.Center },
            MergedCellsList = new List<MergedCells>
            {
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("C", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("C", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("D", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("D", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("E", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("E", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("F", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("F", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("G", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("G", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("H", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("H", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("I", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("I", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("J", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("J", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("K", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("K", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("L", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("L", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                }
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
                        table1St,
                        table2Nd,
                        tableBottomBlueHeader,
                        tableSalaries,
                        tableSharingBeforeAfterData
                    },

                    SheetRows = new()
                    {
                        rowFirstTableHeader,
                        rowSecondTableHeader
                    },

                    SheetCells = new()
                }
            }
        };
    }
}