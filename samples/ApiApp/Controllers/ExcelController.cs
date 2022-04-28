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

    public ExcelController(IExcelWizardService excelWizardService)
    {
        _excelWizardService = excelWizardService;
    }

    [HttpGet("export-excel-from-easy-excel-model")]
    public IActionResult ExportExcelFromExcelWizardModel()
    {
        var excelWizardModel = new CompoundExcelBuilder
        {
            // FileName = "From-Model",

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
                                        Cells = new List<Cell>
                                        {
                                            new(new CellLocation(3,5))
                                            {
                                                Value = "احمد",
                                                CellType= CellType.Text,
                                                CellTextAlign = TextAlign.Center
                                            }
                                        },
                                        MergedCellsList = new(){"C5:D5"},
                                        //StartLocation = new Location(3,5),
                                        //EndLocation = new Location(4,5),
                                        Font = new TextFont{FontColor = Color.DarkGreen},
                                        BackgroundColor = Color.Aqua,
                                        OutsideBorder = new Border
                                        {
                                            BorderLineStyle = LineStyle.DashDotDot,
                                            BorderColor = Color.Brown
                                        }
                                    },
                                    new()
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new(new CellLocation(3,6))
                                            {
                                                Value = "کامبیز دیرباز",
                                                CellType = CellType.Text,
                                                CellTextAlign = TextAlign.Center
                                            }
                                        },
                                        MergedCellsList = new(){"C6:D6"},
                                        //StartLocation = new Location(3,6),
                                        //EndLocation = new Location(4,6),
                                        Font = new TextFont{FontColor = Color.DarkGreen},
                                        BackgroundColor = Color.Aqua,
                                        OutsideBorder = new Border
                                        {
                                            BorderLineStyle = LineStyle.DashDotDot,
                                            BorderColor = Color.Red
                                        }
                                    },
                                    new()
                                    {
                                        Cells = new List<Cell>
                                        {
                                            new(new CellLocation(3,7))
                                            {
                                                Value = "اصغر فرهادی",
                                                CellType = CellType.Text,
                                                CellTextAlign = TextAlign.Center
                                            }
                                        },
                                        MergedCellsList = new(){"C7:D7"},
                                        //StartLocation = new Location(3,7),
                                        //EndLocation = new Location(4,7),
                                        Font = new TextFont{FontColor = Color.DarkGreen},
                                        BackgroundColor = Color.Aqua,
                                        OutsideBorder = new Border()
                                    }
                                },
                                //StartLocation = new Location(3,5), //TODO: Can't be inferred from First Row StartLocation???
                                //EndLocation = new Location(4,7), //TODO: Can't be inferred from EndLocation of last Row???
                                OutsideBorder = new Border
                                {
                                    BorderLineStyle = LineStyle.Thick,
                                    BorderColor = Color.GreenYellow
                                },
                                MergedCells = new List<string>{ "C5:D6" }
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
                            Cells = new List<Cell>
                            {
                                new(new CellLocation(3,2)) {
                                    Value = "فرشاد",
                                    CellType = CellType.Text,
                                    CellTextAlign = TextAlign.Right
                                }
                            },
                            MergedCellsList = new(){"C2:D2"},
                            //StartLocation = new Location(2,2),
                            //EndLocation = new Location(4,2),
                            Font = new TextFont{FontColor = Color.DarkGreen},
                            BackgroundColor = Color.AliceBlue,
                            OutsideBorder = new Border()
                        }
                    },

                    SheetCells = new List<Cell>
                    {
                        new(new CellLocation("A",1)){
                            Value = 11,
                            CellType = CellType.Percentage,
                            CellTextAlign = TextAlign.Left
                        },
                        new(new CellLocation(2, 1))
                        {
                            Value = 112343,
                            CellType = CellType.Currency
                        },
                        new(new CellLocation("D", 1)) { Value = 112 },
                        new(new CellLocation(1, 2))
                        {
                            Value = 211,
                            CellTextAlign = TextAlign.Center
                        },
                        new(new CellLocation(2, 2)) { Value = 212 }
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
}