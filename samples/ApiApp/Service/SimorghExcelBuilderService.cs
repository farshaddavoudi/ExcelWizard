using ApiApp.SimorghReportModels;
using ExcelWizard.Models;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWSheet;
using ExcelWizard.Models.EWStyles;
using ExcelWizard.Models.EWTable;
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

        // Building Excel Parts: Seeing the Excel template At a glance, we can see it is created of below sections:
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
                        CellBuilder
                            .SetLocation("A", 1)
                            .SetValue("گزارش تست")
                            .SetStyle(new CellStyle
                            {
                                CellTextAlign = TextAlign.Center
                            })
                            .Build()
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
                CellBuilder.SetLocation("A", 3).SetValue("کد حساب").Build(),
                CellBuilder.SetLocation("B", 3).SetValue("بدهکار").Build(),
                CellBuilder.SetLocation("C", 3).SetValue("بستانکار").Build()
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
                    CellBuilder.SetLocation("A", 4).SetValue(item.AccountCode).Build(),
                    CellBuilder.SetLocation("B", 4).SetValue(item.Debit).SetContentType(CellContentType.Currency).Build(),
                    CellBuilder.SetLocation("C", 4).SetValue(item.Credit).SetContentType(CellContentType.Currency).Build()
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
                CellBuilder.SetLocation("A", table1St.GetNextVerticalRowNumberAfterTable()).SetValue("کد حساب").Build(),
                CellBuilder.SetLocation("B", table1St.GetNextVerticalRowNumberAfterTable()).SetValue("بدهکار").Build(),
                CellBuilder.SetLocation("C", table1St.GetNextVerticalRowNumberAfterTable()).SetValue("بستانکار").Build(),
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
                    CellBuilder
                        .SetLocation("A", index + rowSecondTableHeader.GetNextRowNumberAfterRow())
                        .SetValue(item.AccountCode)
                        .Build(),
                    CellBuilder
                        .SetLocation("B", index + rowSecondTableHeader.GetNextRowNumberAfterRow())
                        .SetValue(item.Debit)
                        .SetContentType(CellContentType.Currency)
                        .Build(),
                    CellBuilder
                        .SetLocation("C", index + rowSecondTableHeader.GetNextRowNumberAfterRow())
                        .SetValue(item.Credit)
                        .SetContentType(CellContentType.Currency)
                        .Build()
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
                        CellBuilder.SetLocation("A", table2Nd.GetNextVerticalRowNumberAfterTable()).SetValue("نام حساب").Build(),
                        CellBuilder.SetLocation("B", table2Nd.GetNextVerticalRowNumberAfterTable()).SetValue("کد حساب").Build(),
                        CellBuilder.SetLocation("C", table2Nd.GetNextVerticalRowNumberAfterTable()).SetValue("کارخانه دان-51011").Build(),
                        CellBuilder.SetLocation("D", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("E", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("F", table2Nd.GetNextVerticalRowNumberAfterTable()).SetValue("پرورش پولت-51018").Build(),
                        CellBuilder.SetLocation("G", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("H", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("I", table2Nd.GetNextVerticalRowNumberAfterTable()).SetValue("تخم گزار تجاری-51035").Build(),
                        CellBuilder.SetLocation("J", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("K", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("L", table2Nd.GetNextVerticalRowNumberAfterTable())
                            .SetValue("میانگین")
                            .SetStyle(new CellStyle
                            {
                                BackgroundColor = Color.White,
                                Font = new TextFont { FontColor = Color.Black }
                            })
                            .Build()
                    },
                    RowStyle = new RowStyle { RowHeight = 20 }
                },
                new()
                {
                    RowCells = new List<Cell>
                    {
                        CellBuilder.SetLocation("A", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).Build(),
                        CellBuilder.SetLocation("B", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).Build(),
                        CellBuilder.SetLocation("C", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("قبل از تسهیم").Build(),
                        CellBuilder.SetLocation("D", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("بعد از تسهیم").Build(),
                        CellBuilder.SetLocation("E", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("جمع").Build(),
                        CellBuilder.SetLocation("F", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("قبل از تسهیم").Build(),
                        CellBuilder.SetLocation("G", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("بعد از تسهیم").Build(),
                        CellBuilder.SetLocation("H", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("جمع").Build(),
                        CellBuilder.SetLocation("I", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("قبل از تسهیم").Build(),
                        CellBuilder.SetLocation("J", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("بعد از تسهیم").Build(),
                        CellBuilder.SetLocation("K", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("جمع").Build(),
                        CellBuilder
                            .SetLocation("L", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetStyle(new CellStyle
                            {
                                BackgroundColor = Color.White
                            }
                        ).Build()
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
                    CellBuilder.SetLocation("A", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + index).SetValue(account.Name).Build(),
                    CellBuilder.SetLocation("B", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + index).SetValue(account.Code).Build()
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
                        CellBuilder.SetLocation("C", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("D", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("E", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("F", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("G", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("H", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("I", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("J", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("K", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build(),
                        CellBuilder.SetLocation("L", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable()).SetValue(504000).Build()
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

            AllSheetsDefaultStyle = new AllSheetsDefaultStyle
            {
                AllSheetsDefaultTextAlign = TextAlign.Right,
                AllSheetsDefaultDirection = SheetDirection.RightToLeft
            },

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