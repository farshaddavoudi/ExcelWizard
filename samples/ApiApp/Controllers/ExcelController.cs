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
        // For demo, we use static data
        var reportData = new AccountsReportDto
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
            //TODO: Continute building model
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