using System.Drawing;
using ApiApp.DocExampleModels;
using ApiApp.Service;
using ApiApp.SimorghReportModels;
using ExcelWizard.Models;
using ExcelWizard.Service;
using Microsoft.AspNetCore.Mvc;

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

    [HttpGet("export-compound-excel")]
    public IActionResult ExportExcelFromExcelWizardModel()
    {
        // Fetch data from db
        // Here we do not care about the properties and business. Our focus is merely on the Excel report generation
        #region Fetch data from your app business Service

        // For demo, we use static data

        var accountsReportDto = new AccountsReportDto
        {
            ReportName = "Accounts Report",

            AccountDebitCreditList = new List<AccountDebitCredit>
            {
                new() { AccountCode = "13351", Debit = 0, Credit = 50000 },
                new() { AccountCode = "21253", Debit = 50000, Credit = 0 },
                new() { AccountCode = "13556", Debit = 0, Credit = 1000000 },
                new() { AccountCode = "13500", Debit = 0, Credit = 1000000 },
                new() { AccountCode = "13499", Debit = 0, Credit = 2000000 },
                new() { AccountCode = "22500", Debit = 4000000, Credit = 0}
            },

            AccountSalaryCodes = new List<AccountSalaryCode>
            {
                new() { Code = "81010", Name = "Base Salari" },
                new() { Code = "81011", Name = "Overtime Salari" }
            },

            AccountSharingData = new List<AccountSharingData>
            {
                new()
                {
                    AccountName = "Branch 1",
                    AccountSharingDetail = new AccountSharingDetail
                    {
                        BeforeSharing = 504000,
                        AfterSharing = 51353
                    }
                },
                new()
                {
                    AccountName = "Branch 2",
                    AccountSharingDetail = new AccountSharingDetail
                    {
                        BeforeSharing = 11000,
                        AfterSharing = 10000
                    }
                }
            },

            Average = 32000
        };

        #endregion

        // Steps to generate Excel
        // It is the heart of using the ExcelWizard package to generate your desired Excel report
        //---------------------------------------------------------------
        // 1- Analyze Excel Template and Divide It into Separate Sections
        //---------------------------------------------------------------
        //1.1- Top header which is a Table (is not a Row because of occupying two Rows i.e. RowNumber 1 and RowNumber 2) which is Merged and became a Unit Cell
        //1.2 - Having a Row which is the debits credits table Header(It can be part of the debits credits Table model, but it makes it a little hard because the Table data is dynamic and it is better to see the Table header as a Single Row.
        //1.3 - First table with some dynamic data(debits and credits) which the data is in currency type
        //1.4 - Now it is the interesting part! the way I like to see it is a big Table from A10 until I11.There are multiple merges can be seen here, including:
        //A10:A11(Account Name)
        //B10: B11(Account Code)
        //C10: E10(Branch 1)
        //F10: H10(Branch 2)
        //I10: I11(Average)
        //1.5 - Bottom Table with thin inside borders having Base Salary and Overtime Salary Data in it.
        //1.6 - Table with Sharing data which is merged vertically. It can not be considered as Row because, again, being merged and therefore, occupying more than one row.
        //1.7 - A Row with Reporting datetime info
        //1.8 - A Cell with my name on it! at the bottom of Excel
        //-------------------------------------
        //2 - Create each Section Related Model
        //-------------------------------------
        //2.1- Table: Top Header
        var tableTopHeader = new Table
        {
            TableRows = new List<Row>
            {
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("A", 1, accountsReportDto.ReportName) {
                            CellStyle = new CellStyle
                            {
                                // The Cell TextAlign can be set with below property, but because most of the
                                // Cells are TextAlign center, the better approach is to set the Sheet default TextAlign
                                // to Center
                                CellTextAlign = TextAlign.Center
                            }
                        }
                    }
                }
            },
            //TableStyle = new(), //This table do not have any special styles
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
        //2.2- Row: Gray bg row (table Header)
        var rowCreditsDebitsTableHeader = new Row
        {
            RowCells = new List<Cell>
            {
                new("A", 3, "Account Code"),
                new("B", 3, "Debit"),
                new("C", 3, "Credit")
            },

            RowStyle = new RowStyle
            {
                BackgroundColor = Color.LightGray
            }
        };
        //2.3- Table: Credits, Debits table data
        var tableCreditsDebitsData = new Table
        {
            // Using below format is recommended and make it easy to use Collection data and make dynamic Tables/Rows/Cells
            // SomeList.Select((item, index) => ...); item: is an item of collection / index: is the loop index
            TableRows = accountsReportDto.AccountDebitCreditList.Select((item, index) => new Row
            {
                RowCells = new List<Cell>
                {
                    // Notice in getting the Table RowNumber using its top Section (rowCreditsDebitsTableHeader)
                    // You can see this pattern through the rest of codes
                    // So that is the reason building these elements should be step by step and from top to bottom (Remember the Excel data is dynamic and the number of Credits/Debits rows can be varying according to DTO)
                    new("A", rowCreditsDebitsTableHeader.GetNextRowNumberAfterRow() + index, item.AccountCode),
                    new("B", rowCreditsDebitsTableHeader.GetNextRowNumberAfterRow() + index, item.Debit) { CellContentType = CellContentType.Currency },
                    new("C", rowCreditsDebitsTableHeader.GetNextRowNumberAfterRow() + index, item.Credit) { CellContentType = CellContentType.Currency }
                }
            }).ToList(),

            TableStyle = new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick }
            }
        };



        var excelWizardModel = new CompoundExcelBuilder();
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
                new()
                {
                    Name="حقوق پایه",
                    Code="81010"
                },

                new()
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