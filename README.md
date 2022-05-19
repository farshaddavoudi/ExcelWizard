# ExcelWizard
[![NuGet Version](https://img.shields.io/nuget/v/ExcelWizard.svg?style=flat)](https://www.nuget.org/packages/ExcelWizard/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://raw.githubusercontent.com/farshaddavoudi/ExcelWizard/master/LICENSE)

Using ExcelWizard, you can easily generate Excel file in a very simple and straightforward way. In addition, make the generated Excel file directly downloadable from Browser without any hassle in case of using Blazor application. The package is a wrapper for ClosedXML and BlazorFileDownload packages.

# How to Use
1. Install `ExcelWizard` pacakge from nuget package manager.

2. Register ExcelWizard Service in your application (API or Blazor) by using the **AddExcelWizardServices** extension.
```csharp
// Has a `isBlazorApp` argument (default is `false`). 
// In case of using in Blazor application, pass the true value to register necessary services.
// Has an optional argument for ServiceLifeTime. The default lifetime is Scoped.
builder.Services.AddExcelWizardServices(isBlazorApp: false);
```

3. Inject `IExcelWizardService` into your class and enjoy it!

# How much is it simple to generate/download Excel with ExcelWizard?

Assuming you have a collection of a model e.g. `var users = List<User>()` which you normally get it from a database and show them in a Grid or something. Now, 
you decided to have a live Excel report from it anytime you want. Something like below:

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-5.png">

The Model:

```csharp
// The model you want to have Excel report upon it
public class User
{
    public int Id { get; set; }
    public string? FullName { get; set; }
    public string? PersonnelCode { get; set; }
    public string? Nationality { get; set; }
}
```

In your Service or Controller:

```csharp
public class UserController : ControllerBase 
{
    // Inject the IExcelWizardService Service 
    private IExcelWizardService _excelWizardService;

    public ExcelController(IExcelWizardService excelWizardService)
    {
        _excelWizardService = excelWizardService;
    }

    [HttpGet("export-users")]
    public IActionResult ExportUsersToExcel()
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

```

In case you are coding in the Blazor application, the scenario is even simpler. Only get the raw data (=`myUsers`) from API and use the `BlazorDownloadGridLayoutExcel` method
of `ExcelWizardService`, the Excel file will be instantly downloaded (by opening the download popup) from the browser without any struggle for byte[] handling or something :)

In IndexPage.razor:

```razor
<button @onclick="DownloadExcelReport"> Export Excel </button>

@code {

    // Inject the Service 
    [Inject] private IExcelWizardService ExcelWizardService { get; set; } = default!;

    private async Task<DownloadFileResult> DownloadExcelReport()
    {
        // Get your data from API usually by Http call
        var myUsers = await apiService.GetMyUsers();

        // Just pass the data to method and you are good to go ;)
        // This method has an optional parameter `generatedFileName` which is obvious by the name
        return ExcelWizardService.BlazorDownloadGridLayoutExcel(myUsers);
    }
}
```

# Concepts

### The Excel you want can be two types: 

1. **Grid-layout** like data; meaning you have a list of data (again like `myUsers`) and you want to easily export it to Excel. The Excel would be 
relatively simple, having a table-like layout, a header, and data. The first examples in the doc were from this type.

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-3.png">

2. **Compound Excel**; a little more complex than the previous Grid-layout one. This Excel type can include some different Rows, Tables, and special Cells each placed
in different Excel locations. The first type is easier and most straightforward and this type has a different Excel build scenario (Using `GenerateCompoundExcel` method of `IExcelWizardService`).

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-4.png">

### Also, you can have different scenario in saving or retrieving generated Excel file:

1. Get the *`byte[]`* of the Excel file and use it for your use case, e.g. sending to another client to be shown or saving in a database, etc.

2. Save the Excel directly *on disk* and get the full path address to send to the app client or save it in the database.

3. (Blazor app) Normally you want to *show the Excel to the user as exported file and do not want to save it somewhere*. If your app client is 
something other than Blazor (e.g. React, Angular, or MVC), your only choice is to work with generated Excel byte[] data and handle it for 
the result you want, but for Blazor apps the story is very simple. Just use the `BlazorDownloadGridLayoutExcel` and `BlazorDownloadCompoundExcel` methods *(notice the methods name start with BlazorDownload)* from `IExcelWizardService` in some click event and the Excel file will be generated and instantly downloaded *(by opening the download popup)* right from 
the browser. Easy-peasy, huh! :)

Knowing these simple concepts,  you will easily understand the IExcelWizardService methods and will be able to generate your favorable Excel
in a very easy and fast way.

**`IExcelWizardService` all Methods at a glance:**
```csharp
public interface IExcelWizardService
{
    // Generate Simple Grid Excel file from special model configured options with ExcelWizard attributes
    public GeneratedExcelFile GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    // Generate Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes
    public GeneratedExcelFile GenerateGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    // Generate Grid-Layout Excel having multiple Sheets from special model configured options with ExcelWizard attributes
    public string GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder, string savePath);

    // Generate Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes
    public string GenerateGridLayoutExcel(object singleSheetDataList, string savePath, string generatedFileName);

    // Generate Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location
    public GeneratedExcelFile GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    // Generate Excel file and save it in path and return the saved url
    public string GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder, string savePath);

    #region Blazor Application

    // [Blazor only] Generate and Download instantly from Browser the Simple Multiple Sheet Grid Excel file from special model configured options with ExcelWizard attributes in Blazor apps
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    // [Blazor only] Generate and Download instantly from Browser the Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes in Blazor apps
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    // [Blazor only] Generate and Download instantly from Browser the Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location in Blazor apps
    public Task<DownloadFileResult> BlazorDownloadCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    #endregion
}
```

# Generate *Single* Sheet Grid-Layout Excel In Details

For single sheet Grid-layout Excel, it is as easy as passing the data (collection of a model like `var myUsers = new List<User>();`, remember?)
to the `GenerateGridLayoutExcel` method (or `BlazorDownloadGridLayoutExcel` in the case of the Blazor app). It will generate (download in Blazor) a very
simple Excel filled with data without any Excel customization. 

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-5.png" />

You saw a simple example in <i>how much is it simple section</i>. Below we are just adding a variable (`IsLoading`) to show some loading during fetching data and generating Excel:

```razor
<button @onclick="DownloadExcelReport"> Export Excel </button>

@code {

    // Inject the Service 
    [Inject] private IExcelWizardService ExcelWizardService { get; set; } = default!;

    private async Task<DownloadFileResult> DownloadExcelReport()
    {
        // IsLoading is a hypothetical variable to show some Loading in your page
        IsLoading = true;
        
        try {
            // Get your data from API usually by Http call
            var myUsers = await apiService.GetMyUsers();

            // Just pass the data to method and you are good to go ;)
            // This method has an optional parameter `generatedFileName` which is obvious by the name
            return await ExcelWizardService.BlazorDownloadGridLayoutExcel(myUsers);
        }
        finally {
            // Finish Showing loading 
            IsLoading = false;
        }
    }
}
```

### Customize Excel using **`[ExcelSheet]`** and **`[ExcelColumn]`** Attributes
For example, ignore a column (property) to be shown in exported Excel, having some aligns for header or cells, text font/size/color, different background color for header or cells or a specific column!,
custom header name for a column, custom header height or column width or Sheet direction (RTL/LTR), etc..? All these options plus a lot more can be configured by two
attributes you can use on your model. `[ExcelSheet]` for Excel generic properties and `[ExcelColumn]` for per property (column) customization.

Remember the User model, we can use the attributes below to customize our Users Excel:

```csharp
[ExcelSheet(SheetName = "MyUsers", DefaultTextAlign = TextAlign.Center, HeaderBackgroundColor = KnownColor.LightBlue, HeaderHeight = 40,
    BorderType = LineStyle.DashDotDot, DataBackgroundColor = KnownColor.Bisque, DataRowHeight = 25, IsSheetLocked = true,
    SheetDirection = SheetDirection.RightToLeft, FontColor = KnownColor.Red)]
public class User
{
    [ExcelColumn(HeaderName = "UserId", HeaderTextAlign = TextAlign.Right, DataTextAlign = TextAlign.Right, FontColor = KnownColor.Blue)]
    public int Id { get; set; }

    [ExcelColumn(HeaderName = "Name", HeaderTextAlign = TextAlign.Left, FontWeight = FontWeight.Bold)]
    public string? FullName { get; set; }

    [ExcelColumn(HeaderName = "Personnel No", HeaderTextAlign = TextAlign.Left, ColumnWidth = 50, FontSize = 15)]
    public string? PersonnelCode { get; set; }

    public string? Nationality { get; set; }
    
    [ExcelColumn(Ignore = true)]
    public string? Age { get; set; }
}
```

The Result:
<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-2.png" />

You do not need to remember all the properties. Just use the attribute and intellisense will show you all the available options you 
can use to customize Excel.

# Generate *Multiple* Sheets Grid-Layout Excel In Details

It is almost the same with a single sheet, using the second overload of the same `GenerateGridLayoutExcel` method (`BlazorDownloadGridLayoutExcel` method in case of the Blazor app) with the `GridLayoutExcelBuilder` argument 
 that should be provided to configure the Excel file. The Excel customization 
is exactly like Single sheet Grid-layout Excel (See the previous section).

# Generate *Compound* Excel
Generating Excel in this case for single or multi sheets are the same. Using the `GenerateCompoundExcel` method (`BlazorDownloadCompoundExcel` in case of Blazor application) you can create any customized Excel file. Just go along with the `CompoundExcelBuilder` argument and provide the necessary parts for your Excel.

Tip: We do not use any attributes (`[ExcelSheet]` and `[ExcelColumn]`) here.

### Complete Example of Building Compound Excel from Scratch
Let's assume we have an application related to a company's financials and we want to have a custom Excel report like the below format:

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/accounts-excel-template-report.png" />

You can download Excel from <a target="_blank" href="https://github.com/farshaddavoudi/ExcelWizard/blob/main/templates/CompoundExcelTemplate.xlsx">here</a>.

So, it is a **Compound Excel**, not the **GridLayout** one. 

We should create our Excel template (`CompoundExcelBuilder` model) using our local app model (here `accountsReportDto`), So At first, we should fetch the report model (DTO) (normally from a database).

In our example, the data is like below:

```csharp
        // Fetch data from db
        // For demo, we use static data
        // Here we do not care about the properties and business. Our focus is merely on the Excel report generation
        // All Classes can be seen in Repo Samples
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
```

Now that we have our report data, we can create our Excel step by step.

### Steps to Generate the Compound Excel Filled with DTO data: 

**1- Analyze Excel Template and Divide It into Separate Sections <br />
2- Create each Section Related Model <br />
3- Create `CompoundExcelBuilder` Model <br />
4- Generate Excel using `ExcelWizardService` and `CompoundExcelBuilder` model (step 3)**

## *1- Analyze Excel Template and Divide It into Separate Sections*

Analyze the Excel template and divide it into **Table**s, **Row**s, and **Cell**s sections. In the next step, each section will be mapped to its ExcelWizard model equivalent. We use these section models in Step 3 to create the `CompoundExcelBuilder` model.

For our example, seeing the Excel template at a glance, we can detect it is composed of below sections:

1- Top header which is a **Table** (is not a **Row** because of occupying two Rows *i.e. RowNumber 1 and RowNumber 2*) which is Merged and became a Unit Cell

2- Having a **Row** which is the debits credits table Header (It can be part of the debits credits **Table** model, but it makes it a little hard because
   the Table data is dynamic and it is better to see the Table header as a Single Row.

3- First table with some dynamic data (debits and credits) which the data is in currency type

4- Now it is the interesting part! the way I like to see it is a big **Table** from **A10** until **I11**. There are 
multiple merges can be seen here, including:
- `A10:A11 (Account Name)` 
- `B10:B11 (Account Code)` 
- `C10:E10 (Branch 1)`  
- `F10:H10 (Branch 2)`
- `I10:I11 (Average)`

5- Bottom **Table** with thin inside borders having *Base Salary* and *Overtime Salary* Data in it.

6- **Table** with *Sharing data* which is merged vertically. It can not be considered as Row because, again, being merged and therefore, occupying more than one row.

7- A **Row** with Reporting datetime info

8- A **Cell** with my name on it! at the bottom of Excel


## *2- Create each Section Related Model*

These models are `Table` model, `Row` model and `Cell` model. All of them are ExcelWizard models and will be used in generating the main `CompoundExcelBuilder` model (in the next step).
Note in creating these models that, all properties have proper comments to make them clear and their names also speak for themselves.

**1- Table: Top Header**
```csharp
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
```

**2- Row: Gray bg row (table Header)**
```csharp
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
```

**3- Table: Credits, Debits table data**
```csharp
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
                    // So that is the reason building these elements should be step by step and from top to bottom 
                    // Remember the Excel data is dynamic and the number of Credits/Debits rows can be varying according to DTO, so the row counts cannot be static
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
```

**4- Table: Blue bg (+yellow at the end) table**
```csharp
var tableBlueBg = new Table
        {
            TableRows = new List<Row>
            {
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("A", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable(), "Account Name"),
                        new("B", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable(), "Account Code"),
                        new("C", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable(), "Branch 1"),
                        new("D", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        new("E", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        new("F", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable(), "Branch 2"),
                        new("G", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        new("H", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        new("I", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable(), "Average")
                        {
                            CellStyle =
                            {
                                //BackgroundColor = Color.Yellow, //Bg will set on Merged properties
                                Font = new TextFont { FontColor = Color.Black }
                            }
                        }
                    },
                    RowStyle = new RowStyle { RowHeight = 20 }
                },
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("A", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1),
                        new("B", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1),
                        new("C", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1, "Before Sharing"),
                        new("D", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1, "After Sharing"),
                        new("E", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1, "Sum"),
                        new("F", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1, "Before Sharing"),
                        new("G", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1, "After Sharing"),
                        new("H", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1, "Sum"),
                        new("I", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1)
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
                        FirstCellLocation = new CellLocation("A", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("A", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("B", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("B", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("C", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("E", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable())
                    },
                    BackgroundColor = Color.DarkBlue

                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("F", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("H", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable())
                    },
                    BackgroundColor = Color.DarkBlue
                }
                ,
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("I", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("I", tableCreditsDebitsData.GetNextVerticalRowNumberAfterTable() + 1)
                    },
                    BackgroundColor = Color.Yellow
                }
            }
        };
```

**5- Table: with Salaries data with thin borders**
```csharp
var tableSalaries = new Table
        {
            TableRows = accountsReportDto.AccountSalaryCodes.Select((account, index) => new Row
            {
                RowCells = new List<Cell>
                {
                    new ("A", tableBlueBg.GetNextVerticalRowNumberAfterTable() + index, account.Name),
                    new ("B", tableBlueBg.GetNextVerticalRowNumberAfterTable() + index, account.Code)
                }
            }).ToList(),
            TableStyle = new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black }
            }
        };
```
**6- Table:  Sharing info**
Table with sharing before/after data
```csharp
var tableSharingBeforeAfterData = new Table
        {
            TableRows = new List<Row>
            {
                new()
                {
                    RowCells = new List<Cell>
                    {
                        new("C", tableBlueBg.GetNextVerticalRowNumberAfterTable(), accountsReportDto.AccountSharingData
                            .Where(s => s.AccountName == "Branch 1")
                            .Select(s => s.AccountSharingDetail.BeforeSharing)
                            .FirstOrDefault()),
                        new("D", tableBlueBg.GetNextVerticalRowNumberAfterTable(), accountsReportDto.AccountSharingData
                            .Where(s => s.AccountName == "Branch 1")
                            .Select(s => s.AccountSharingDetail.AfterSharing)
                            .FirstOrDefault()),
                        new("E", tableBlueBg.GetNextVerticalRowNumberAfterTable(), accountsReportDto.AccountSharingData
                            .Where(s => s.AccountName == "Branch 1")
                            .Select(s => s.AccountSharingDetail.AfterSharing + s.AccountSharingDetail.BeforeSharing)
                            .FirstOrDefault()),
                        new("F", tableBlueBg.GetNextVerticalRowNumberAfterTable(), 11000),
                        new("G", tableBlueBg.GetNextVerticalRowNumberAfterTable(), 10000),
                        new("H", tableBlueBg.GetNextVerticalRowNumberAfterTable(), 21000),
                        new("I", tableBlueBg.GetNextVerticalRowNumberAfterTable(), accountsReportDto.Average)
                    }
                }
            },
            //TableStyle = new TableStyle { TableTextAlign = TextAlign.Center }, //Inherit from Sheet TextAlign Center
            MergedCellsList = new List<MergedCells>
            {
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("C", tableBlueBg.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("C", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("D", tableBlueBg.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("D", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("E", tableBlueBg.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("E", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("F", tableBlueBg.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("F", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("G", tableBlueBg.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("G", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("H", tableBlueBg.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("H", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                },
                new()
                {
                    MergedBoundaryLocation = new MergedBoundaryLocation
                    {
                        FirstCellLocation = new CellLocation("I", tableBlueBg.GetNextVerticalRowNumberAfterTable()),
                        LastCellLocation = new CellLocation("I", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    }
                }
            }
        };
```

**7- Row: Light Green row for report date**

```csharp
var rowReportDate = new Row
        {
            RowCells = new List<Cell>
            {
                new Cell("D", tableSharingBeforeAfterData.GetNextVerticalRowNumberAfterTable() + 1, DateTime.Now)
            },
            MergedCellsList = new List<MergedBoundaryLocation>
            {
                new()
                {
                    FirstCellLocation = new CellLocation("D", tableSharingBeforeAfterData.GetNextVerticalRowNumberAfterTable() + 1),
                    LastCellLocation = new CellLocation("F", tableSharingBeforeAfterData.GetNextVerticalRowNumberAfterTable() + 1)
                }
            }
        };
```

**8- Cell: User name (me!)**

```csharp
var cellUserName = new Cell("E", rowReportDate.GetNextRowNumberAfterRow() + 1, "Farshad Davoudi")
        {
            CellStyle = new CellStyle
            {
                BackgroundColor = Color.DarkGreen,
                Font = new TextFont
                {
                    FontColor = Color.White
                },
                CellBorder = new Border
                {
                    BorderLineStyle = LineStyle.Thin,
                    BorderColor = Color.Red
                }
            }
        };
```

## *3- Create `CompoundExcelBuilder` Model*

Then we create our main model by using the sections model created in Step 2 plus other styles that are available in this class.

```csharp
var excelWizardModel = new CompoundExcelBuilder
        {
            GeneratedFileName = "AccountsReport",
            AllSheetsDefaultStyle = new AllSheetsDefaultStyle
            {
                AllSheetsDefaultTextAlign = TextAlign.Center
            },
            Sheets = new List<Sheet>
            {
                new Sheet()
                {
                    SheetTables = new List<Table>
                    {
                        tableTopHeader,

                        tableCreditsDebitsData,

                        tableBlueBg,

                        tableSalaries,

                        tableSharingBeforeAfterData
                    },

                    SheetRows = new List<Row>
                    {
                        rowCreditsDebitsTableHeader,

                        rowReportDate
                    },

                    SheetCells = new List<Cell>
                    {
                        cellUserName
                    }
                }
            }
        };
```

## *4- Generate Excel using `ExcelWizardService` and `CompoundExcelBuilder` model (step 3)*

At last, we create our gorgeous Excel! by injecting `IExcelWizardService` and using the Step 3 `compoundExcelBuilder` model. It is the easiest part! 

```csharp
return Ok(_excelWizardService.GenerateCompoundExcel(excelWizardModel, @"C:\GeneratedExcelSamples"));
```
