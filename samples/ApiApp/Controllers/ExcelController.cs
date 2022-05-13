using ApiApp.Service;
using ApiApp.SimorghReportModels;
using ExcelWizard.Models;
using ExcelWizard.Service;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;

namespace ApiApp.Controllers;

[Route("api/[controller]")]
[ApiController]
public class ExcelController : ControllerBase
{
    private readonly IExcelWizardService _excelWizardService;
    private readonly ISimorghExcelBuilderService _simorghExcelBuilderService;

    public ExcelController(IExcelWizardService excelWizardService, ISimorghExcelBuilderService simorghExcelBuilderService)
    {
        _excelWizardService = excelWizardService;
        _simorghExcelBuilderService = simorghExcelBuilderService;
    }


    [HttpGet("export-simple-compound-excel")]
    public IActionResult ExportExcelFromExcelWizardModel()
    {
        var excelWizardModel = new CompoundExcelBuilder
        {
            // GeneratedFileName = "From-Model",

            Sheets = new List<Sheet>
            {
                new Sheet
                {
                    SheetTables = new List<Table>
                    {
                        new()
                            {
                                TableRows = new List<Row>
                                {
                                    new()
                                    {
                                        RowCells = new List<Cell>
                                        {
                                            new(3,5)
                                            {
                                                Value = "احمد",
                                                CellContentType = CellContentType.Text,
                                                CellStyle = new CellStyle
                                                {
                                                    CellTextAlign = TextAlign.Center
                                                }
                                            }
                                        },
                                        MergedCellsList = new()
                                        {
                                            new MergedBoundaryLocation
                                            {
                                                FirstCellLocation = new CellLocation("C", 5),
                                                LastCellLocation = new CellLocation("D", 5)
                                            }
                                        },
                                        RowStyle = new RowStyle
                                        {
                                            Font = new TextFont{FontColor = Color.DarkGreen},
                                            BackgroundColor = Color.Aqua,
                                            OutsideBorder = new Border
                                            {
                                                BorderLineStyle = LineStyle.DashDotDot,
                                                BorderColor = Color.Brown
                                            }
                                        }
                                    },
                                    new()
                                    {
                                        RowCells = new List<Cell>
                                        {
                                            new(3,6)
                                            {
                                                Value = "کامبیز دیرباز",
                                                CellContentType = CellContentType.Text,
                                                CellStyle = new CellStyle
                                                {
                                                    CellTextAlign = TextAlign.Center
                                                }
                                            }
                                        },
                                        MergedCellsList = new()
                                        {
                                            new MergedBoundaryLocation
                                            {
                                                FirstCellLocation = new CellLocation("C", 6),
                                                LastCellLocation = new CellLocation("D", 6)
                                            }
                                        },
                                        RowStyle = new RowStyle
                                        {
                                            Font = new TextFont{FontColor = Color.DarkGreen},
                                            BackgroundColor = Color.Aqua,
                                            OutsideBorder = new Border
                                            {
                                                BorderLineStyle = LineStyle.DashDotDot,
                                                BorderColor = Color.Red
                                            }
                                        }
                                    },
                                    new()
                                    {
                                        RowCells = new List<Cell>
                                        {
                                            new(3,7)
                                            {
                                                Value = "اصغر فرهادی",
                                                CellContentType = CellContentType.Text
                                            }
                                        },
                                        MergedCellsList = new()
                                        {
                                            new MergedBoundaryLocation
                                            {
                                                FirstCellLocation = new CellLocation("C", 7),
                                                LastCellLocation = new CellLocation("D", 7)
                                            }
                                        },
                                        RowStyle = new RowStyle
                                        {
                                            Font = new TextFont{FontColor = Color.DarkGreen},
                                            BackgroundColor = Color.Aqua,
                                            OutsideBorder = new Border()
                                        }
                                    }
                                },
                                //StartLocation = new Location(3,5), //TODO: Can't be inferred from First Row StartLocation???
                                //EndLocation = new Location(4,7), //TODO: Can't be inferred from EndLocation of last Row???
                                TableStyle = new TableStyle
                                {
                                    OutsideBorder = new Border
                                    {
                                        BorderLineStyle = LineStyle.Thick,
                                        BorderColor = Color.GreenYellow
                                    }
                                },
                                MergedCellsList = new List<MergedCells>
                                {
                                    new()
                                    {
                                        MergedBoundaryLocation = new()
                                        {
                                            FirstCellLocation = new CellLocation("C", 5),
                                            LastCellLocation = new CellLocation("D", 6)
                                        }
                                    }
                                }
                            }
                    },

                    SheetColumnsStyle = new List<ColumnStyle>
                    {
                        new() { ColumnNo = 3, ColumnWidth = new ColumnWidth{ Width = 30 } },
                        new() { ColumnNo = 1, IsColumnLocked = true, ColumnWidth = new ColumnWidth{ WidthCalculationType = ColumnWidthCalculationType.AdjustToContents }}
                    },

                    SheetRows = new List<Row>
                    {
                        new()
                        {
                            RowCells = new List<Cell>
                            {
                                new(3,2) {
                                    Value = "فرشاد",
                                    CellContentType = CellContentType.Text,
                                    CellStyle = new CellStyle
                                    {
                                        CellTextAlign = TextAlign.Right
                                    }
                                }
                            },
                            MergedCellsList = new()
                            {
                                new MergedBoundaryLocation
                                {
                                    FirstCellLocation = new CellLocation("C", 2),
                                    LastCellLocation = new CellLocation("D", 2)
                                }
                            },
                            RowStyle = new RowStyle
                            {
                                Font = new TextFont{FontColor = Color.DarkGreen},
                                BackgroundColor = Color.AliceBlue,
                                OutsideBorder = new Border()
                            }
                        }
                    },

                    SheetCells = new List<Cell>
                    {
                        new("A",1){
                            Value = 11,
                            CellContentType = CellContentType.Percentage,
                            CellStyle = new CellStyle
                            {
                                CellTextAlign = TextAlign.Left
                            }
                        },
                        new(2, 1)
                        {
                            Value = 112343,
                            CellContentType = CellContentType.Currency
                        },
                        new("D", 1) { Value = 112 },
                        new(1, 2)
                        {
                            Value = 211,
                            CellStyle = new CellStyle
                            {
                                CellTextAlign = TextAlign.Center
                            }
                        },
                        new(2, 2) { Value = 212 }
                    }
                }
            }
        };

        return Ok(_excelWizardService.GenerateCompoundExcel(excelWizardModel, @"C:\GeneratedExcelSamples"));
    }

    [HttpGet("export-grid-excel")]
    public IActionResult ExportGridExcel()
    {
        // The below data normally comes from your database
        // Show static for demo purposes
        var myUsers = new List<User>
        {
            new() { Id = 1, FullName = "Ronaldo", PersonnelCode = "980923", Nationality = "Portugal" },
            new() { Id = 2, FullName = "Messi", PersonnelCode = "991126", Nationality = "Argentine" },
            new() { Id = 3, FullName = "Mbappe", PersonnelCode = "991213", Nationality = "France" }
        };

        // Below will create Excel file as byte[] data
        // Just passing your data to method argument and let the rest to the package! hoorya!
        // This method has an optional parameter `generatedFileName` which is obvious by the name
        GeneratedExcelFile generatedExcelFile = _excelWizardService.GenerateGridLayoutExcel(myUsers);

        // Below will create Excel file in specified path and return the full path as string
        // The last param is generated file name
        string fullPathAsString = _excelWizardService.GenerateGridLayoutExcel(myUsers, @"C:\GeneratedExcelSamples", "Users-Excel");

        return Ok(generatedExcelFile);
    }

    [HttpGet("export-simorgh-report-compound-excel")]
    public IActionResult ExportExcelFromSimorghCompoundReport()
    {
        var voucherStatementPageResult = new VoucherStatementResult
        {
            ReportName = "ExcelWizard Compound Report",

            SummaryAccounts = new List<SummaryAccount>
                {
                    new SummaryAccount
                    {
                        AccountName = "کارخانه دان-51011" ,
                        Multiplex =new List<Multiplex>
                        {
                            new() { After = 5000000, Before = 4000 }
                        }
                    },

                    new SummaryAccount
                    {
                        AccountName = "پرورش پولت-51018" ,
                        Multiplex =new List<Multiplex>
                        {
                            new() { After = 5000000, Before = 4000 }
                        }
                    },

                    new SummaryAccount
                    {
                        AccountName = "تخم گزار تجاری-51035" ,
                        Multiplex =new List<Multiplex>
                        {
                            new Multiplex{After = 5000000,Before = 4000 }
                        }
                    }
                },

            Accounts = new List<AccountDto>
                {
                   new AccountDto
                   {
                        Name="حقوق پایه",
                        Code="81010"
                   },

                   new AccountDto
                   {
                        Name="اضافه کار",
                        Code="81011"
                   }
                },

            VoucherStatementItem = new List<VoucherStatementItem>
                {
                    new VoucherStatementItem
                    {
                        AccountCode = "13351",
                        Credit = 50000,
                        Debit = 0
                    },

                    new VoucherStatementItem
                    {
                        AccountCode = "21253",
                        Credit = 0,
                        Debit = 50000
                    },

                    new VoucherStatementItem
                    {
                        AccountCode = "13556",
                        Credit = 1000000,
                        Debit = 0
                    },

                    new VoucherStatementItem
                    {
                        AccountCode = "13500",
                        Credit = 1000000,
                        Debit = 0
                    },

                    new VoucherStatementItem
                    {
                        AccountCode = "13499",
                        Credit = 2000000,
                        Debit = 0
                    },

                    new VoucherStatementItem
                    {
                        AccountCode = "22500",
                        Credit = 0,
                        Debit = 4000000
                    }
                }
        };

        var excel = _simorghExcelBuilderService.GenerateVoucherStatementExcelReport(voucherStatementPageResult, @"C:\GeneratedExcelSamples");

        return Ok(excel);
    }
}