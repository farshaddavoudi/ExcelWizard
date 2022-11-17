# ExcelWizard
[![NuGet Version](https://img.shields.io/nuget/v/ExcelWizard.svg?style=flat)](https://www.nuget.org/packages/ExcelWizard/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://raw.githubusercontent.com/farshaddavoudi/ExcelWizard/master/LICENSE)

Using ExcelWizard, you can easily generate Excel file in a very simple and straightforward way, even without any previous Excel knowledge. In addition, make the generated Excel file directly downloadable from Browser without any hassle in case of using Blazor application. The package is a wrapper for ClosedXML and BlazorFileDownload packages.

## Version >= 3.0.0 Breakthrough Changes
#### The package has completely rewritten with advanced *builder design pattern* to be more user friendly and easier to use and extremely less complex with some added new features like easily and dynamically create Table component using model binding.

# How to Use
1. You can install the package via the nuget package manager just search for *ExcelWizard*. You can also install via powershell using the following command.

```powershell
Install-Package ExcelWizard
```

Or via the dotnet CLI.

```bash
dotnet add package ExcelWizard
```

2. Register ExcelWizard Service in your application (API or Blazor) by using the **`AddExcelWizardServices`** extension.
```csharp
// Has a `isBlazorApp` argument (default is `false`). 
// In case of using in Blazor application, pass the true value to register necessary services.
// Has an optional argument for ServiceLifeTime. The default lifetime is Scoped.
builder.Services.AddExcelWizardServices(isBlazorApp: false);
```

3. Inject **`IExcelWizardService`** into your class and enjoy it!

### How much is it simple to generate Excel with ExcelWizard?

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
        // This method has a `IExcelBuilder` argument. We just need to provide the argument by using ExcelBuilder
        // The chained methods of ExcelBuilder which are in logical order and have proper comments will guide us through the process of creating the model
        IExcelBuilder excelBuilder = ExcelBuilder
            .SetGeneratedFileName("Users")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingAModelToBind(myUsers)
            .Build();

        GeneratedExcelFile generatedExcelFile = _excelWizardService.GenerateExcel(excelBuilder);

        // Below will create Excel file in specified path and return the full path as string     
        string fullPathAsString = _excelWizardService.GenerateExcel(excelBuilder, @"C:\GeneratedExcelSamples");

        return Ok(generatedExcelFile);
    }
}

```

In case you are coding in the Blazor application, the scenario is even simpler. Only get the raw data (=`myUsers`) from API and use the `GenerateAndDownloadExcelInBlazor` method
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

        IExcelBuilder excelBuilder = ExcelBuilder
            .SetGeneratedFileName("my-custom-report")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingAModelToBind(fetchDataFromApi)
            .Build();

        return ExcelWizardService.GenerateAndDownloadExcelInBlazor(excelBuilder);
    }
}
```

# Concepts

#### The Excel you want to export can be in two layout types; *Grid-Layout* and *Complex-Layout*. 
The reason behind this seperation is the much easier process of generating Excel for Grid-Layout comparing to Complex-Layout because of ability to be created with model binding in this type which is pretty easy. 

1. **Grid-Layout** like data; meaning you have a list of data (again like `myUsers`) and you want to easily export it to Excel. The Excel would be 
relatively simple, having a table-like layout, a header, and data. The first examples in the doc were from this type.

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-3.png">

2. **Complex-Layout**; a little more complex than the previous Grid-layout one. This Excel type can divided to smaller sections including Table(s), Row(s) and Cell(s) each placed
in different Excel locations. The first type is easier and most straightforward.

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-4.png">

### Generated Excel Return Type

#### Also, you can have different scenario in saving or retrieving generated Excel file:

1. Get the *`byte[]`* of the Excel file and use it for your use case, e.g. sending to another client to be shown or saving in a database, etc.

2. Save the Excel directly *on disk* and get the full path address to send to the app client or save it in the database.

3. (Blazor app) Normally you want to *show the Excel to the user as exported file and do not want to save it somewhere*. If your app client is 
something other than Blazor (e.g. React, Angular, or MVC), your only choice is to work with generated Excel byte[] data and handle it for 
the result you want, but for Blazor apps the story is very simple. Just use the `GenerateAndDownloadExcelInBlazor` from `IExcelWizardService` in some click event and the Excel file will be generated and instantly downloaded *(by opening the download popup)* right from 
the browser. Easy-peasy, huh! :)

Knowing these simple concepts,  you will easily understand the IExcelWizardService methods and will be able to generate your favorable Excel
in a very easy and fast way.

**`IExcelWizardService` all methods at a glance:**
```csharp
public interface IExcelWizardService
{
    /// <summary>
    /// Generate Excel file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelBuilder"> ExcelBuilder with Build() method at the end </param>
    /// <returns> Byte array of generated Excel saved in memory. </returns>
    GeneratedExcelFile GenerateExcel(IExcelBuilder excelBuilder);

    /// <summary>
    /// Generate Excel file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelBuilder"> ExcelBuilder with Build() method at the end </param>
    /// <param name="savePath"> The url saved </param>
    /// <returns> Save generated Excel in a path in your device </returns>
    string GenerateExcel(IExcelBuilder excelBuilder, string savePath);

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the generated file by providing equivalent CSharp model
    /// </summary>
    /// <param name="excelBuilder">  ExcelBuilder with Build() method at the end </param>
    /// <returns> Instantly download from Browser </returns>
    Task<DownloadFileResult> GenerateAndDownloadExcelInBlazor(IExcelBuilder excelBuilder);
}
```

# Create Excel In Action

It do not matter which type our target Excel is, either grid-layout or complex-layout, we always start with `ExcelBuilder`
and will be guided step by step by chained methods of builder easily until we call the `Build()` method at the end. All
methods are well documented and commented and will come up in a logical order so it is simpler and easier to use without 
fear of errors or missing something.

## Generate Grid-Layout Excel

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

            var excelBuilder = ExcelBuilder
            .SetGeneratedFileName("my-custom-report")
            .CreateGridLayoutExcel()
            .WithOneSheetUsingAModelToBind(myUsers)
            .Build();

            return ExcelWizardService.GenerateAndDownloadExcelInBlazor(excelBuilder);
        }
        finally {
            // Finish Showing loading 
            IsLoading = false;
        }
    }
}
```

### Customize Excel using **`[ExcelSheet]`** and **`[ExcelSheetColumn]`** Attributes
For example, ignore a column (property) to be shown in exported Excel, having some aligns for header or cells, text font/size/color, different background color for header or cells or a specific column!,
custom header name for a column, custom header height or column width or Sheet direction (RTL/LTR), etc..? All these options plus a lot more can be configured by two
attributes you can use on your model. `[ExcelSheet]` for Excel generic properties and `[ExcelSheetColumn]` for per property (column) customization.

Remember the User model, we can use the attributes below to customize our Users Excel:

```csharp
[ExcelSheet(SheetName = "MyUsers", DefaultTextAlign = TextAlign.Center, HeaderBackgroundColor = KnownColor.LightBlue, HeaderHeight = 40,
    BorderType = LineStyle.DashDotDot, DataBackgroundColor = KnownColor.Bisque, DataRowHeight = 25, IsSheetLocked = true,
    SheetDirection = SheetDirection.RightToLeft, FontColor = KnownColor.Red)]
public class User
{
    [ExcelSheetColumn(HeaderName = "UserId", HeaderTextAlign = TextAlign.Right, DataTextAlign = TextAlign.Right, FontColor = KnownColor.Blue)]
    public int Id { get; set; }

    [ExcelSheetColumn(HeaderName = "Name", HeaderTextAlign = TextAlign.Left, FontWeight = FontWeight.Bold)]
    public string? FullName { get; set; }

    [ExcelSheetColumn(HeaderName = "Personnel No", HeaderTextAlign = TextAlign.Left, ColumnWidth = 50, FontSize = 15)]
    public string? PersonnelCode { get; set; }

    public string? Nationality { get; set; }
    
    [ExcelSheetColumn(Ignore = true)]
    public string? Age { get; set; }
}
```

The Result:
<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-2.png" />

## Generate **Complex-Layout** Excel
You can create any customized Excel file. Just go along with the `ExcelBuilder` chained methods and provide the necessary methods for your Excel.

*Tip*: We **do not** use the Grid-Layout configurable attributes (`[ExcelSheet]` and `[ExcelSheetColumn]`) here.

### Complete Example of Building Complex-Layout Excel from Scratch
Let's assume we have an application related to a company's financials and we want to have a custom Excel report like the below format:

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/accounts-excel-template-report.png" />

You can download Excel from <a target="_blank" href="https://github.com/farshaddavoudi/ExcelWizard/blob/main/templates/ComplexExcelTemplate.xlsx">here</a>.

So, obviously, it is a Complex-Layout Excel, not the Grid-Layout one. 

We want to create our `IExcelBuilder` argument using our dynamic local app model (here `accountsReportDto`), So At first, we should fetch the report model (DTO) (normally from a database).

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

### Steps to Generate the Complex Excel Filled with DTO data: 

**1- Analyze Excel Template and Divide It into Separate Sections (Sub-Components), including Table(s), Row(s) and Cell(s) <br />
2- Create each defined Section Builder, e.g. `ITableBuilder` <br />
3- Create main model necessary to generate Excel, i.e. `IExcelBuilder` (Using `ExcelBuilder` and its chained methods) <br />
4- Generate Excel using `ExcelWizardService` and `IExcelBuilder` model (provided at step 3)**

## *1- Analyze Excel Template and Divide It into Separate Sections*

Analyze the Excel template and divide it into **Table**s, **Row**s, and **Cell**s sections. The priority here is with *Table*, then *Row* and at last *Cell*, meaning if you
can define a section as a Table, don't consider it as multiple Rows! or even a lot of Cells! not that it is impossible
to create the Excel model that way or some error will be thrown, but because it makes it harder for you to create your desired Excel using many small components
rather than one bigger component.

In the next step, each section will be mapped to its Builder equivalent, i.e. **ITableBuilder**, **IRowBuilder** and **ICellBuilder**. 
We use these section models in Step 3 to create the `IExcelBuilder` type.

For our example, seeing the Excel template at a glance, we can detect it is composed of below sections:

1- Top header which is a **Table** (is not a **Row** because of occupying two Rows *i.e. RowNumber 1 and RowNumber 2*) which is Merged and became a Unit Cell,
therefore we create a `ITableBuilder`for it. It has again chained methods with lots of comments which guide you step by step
in logical order to create your table builder.

2- First table with some dynamic data (debits and credits) which the data is in currency type. We use automatic model
binding here to create `ITableBuilder`. We configure all properties like header name, header background, table background and etc in the model itself with the help of `[ExcelTable]` and
`[ExcelTableColumn]` attributes

3- Now it is the interesting part! the way I like to see it is a big **Table** from **A10** until **I11**. There are 
multiple merges can be seen here, including:
- `A10:A11 (Account Name)` 
- `B10:B11 (Account Code)` 
- `C10:E10 (Branch 1)`  
- `F10:H10 (Branch 2)`
- `I10:I11 (Average)`
We create this table manually without binding (this table cannot be bound to any model due to its complex layout)

*Important tip:* It can be seen not this way I consider it. For example one table for Accounts and another for Salaries, etc and it is still completely valid.

4- Bottom **Table** with thin inside borders having *Base Salary* and *Overtime Salary* Data in it.

5- **Table** with *Sharing data* which is merged vertically. It can not be considered as Row because, again, being merged and therefore, occupying more than one row.

6- A **Row** with Reporting datetime info

7- A **Cell** with my name on it! at the bottom of Excel


## *2- Create Each Section Related Builder*

These builders are `ITableBuilder`, `IRowBuilder` and `ICellBuilder`. All of them are a builder and will be used in generating the main `IExcelBuilder` model (in the next step).
Note in creating these builders that all start with their builder name (e.g. `TableBuilder`) and finish with `.Build()` method.
The methods are well commented and more importantly, they are chained in logical order to guide you easily in the process and the mothods names are clear and speak for themselves.

**1- Table: Top Header**
```csharp
var tableTopHeader = TableBuilder
    .CreateStepByStepManually()
    .SetRows(RowBuilder
        .SetCells(
            CellBuilder.SetLocation("A", 1)
                .SetValue(accountsReportDto.ReportName)
                .SetStyle(new CellStyle
                {
                    // The Cell TextAlign can be set with below property, but because most of the
                    // Cells are TextAlign center, the better approach is to set the Sheet default TextAlign
                    // to Center
                    CellTextAlign = TextAlign.Center
                })
                .Build())
        .NoMergedCells()
        .NoCustomStyle()
        .Build())
    .SetTableMergedCells(
        MergeBuilder
            .SetMergingStartPoint("A", 1)
            .SetMergingFinishPoint("H", 2)
            .Build()
    )
    .NoCustomStyle()
    .Build();
```

**2- Table: Credits Debits table with new concept of model binding**
```csharp
ITableBuilder tableCreditsDebits = TableBuilder
            .CreateUsingAModelToBind(accountsReportDto.AccountDebitCreditList, new CellLocation("A", 3))
            .NoMergedCells()
            .Build();
```

By the way the binded model of `AccountDebitCredit` is like below:

```csharp
[ExcelTable(HeaderBackgroundColor = KnownColor.LightGray,
    InsideCellsBorderStyle = LineStyle.Thick,
    InsideCellsBorderColor = KnownColor.Black,
    OutsideBorderColor = KnownColor.Black,
    OutsideBorderStyle = LineStyle.Thick,
    FontColor = KnownColor.Blue,
    HasHeader = true,
    FontSize = 11,
    TextAlign = TextAlign.Center)]
public class AccountDebitCredit
{
    [ExcelTableColumn(HeaderName = "Account Code", FontColor = KnownColor.DarkOrange, DataTextAlign = TextAlign.Right,
        HeaderTextAlign = TextAlign.Left, FontSize = 13, FontWeight = FontWeight.Bold)]
    public string? AccountCode { get; set; }

    [ExcelTableColumn(DataContentType = CellContentType.Currency)]
    public decimal Debit { get; set; }

    [ExcelTableColumn(DataContentType = CellContentType.Currency, Ignore = false)]
    public decimal Credit { get; set; }
}
```

**3- Table: Blue bg (+yellow at the end) table**. It is a fat table by the way!
```csharp
var tableBlueBg = TableBuilder
    .CreateStepByStepManually()
    .SetRows(RowBuilder
            .SetCells(
                CellBuilder.SetLocation("A", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
                    .SetValue("Account Name").Build(),
                CellBuilder.SetLocation("B", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
                    .SetValue("Account Code").Build(),
                CellBuilder.SetLocation("C", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
                    .SetValue("Branch 1").Build(),
                CellBuilder.SetLocation("D", tableCreditsDebits.GetNextVerticalRowNumberAfterTable()).Build(),
                CellBuilder.SetLocation("E", tableCreditsDebits.GetNextVerticalRowNumberAfterTable()).Build(),
                CellBuilder.SetLocation("F", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
                    .SetValue("Branch 2").Build(),
                CellBuilder.SetLocation("G", tableCreditsDebits.GetNextVerticalRowNumberAfterTable()).Build(),
                CellBuilder.SetLocation("H", tableCreditsDebits.GetNextVerticalRowNumberAfterTable()).Build(),
                CellBuilder.SetLocation("I", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
                    .SetValue("Average")
                    .SetStyle(new CellStyle
                    {
                        //BackgroundColor = Color.Yellow, //Bg will set on Merged properties
                        Font = new TextFont { FontColor = Color.Black }
                    })
                    .Build()
            )
            .SetRowMergedCells(MergeBuilder
                .SetMergingStartPoint("C", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
                .SetMergingFinishPoint("E", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
                .SetMergingAreaBackgroundColor(Color.Red)
                .Build())
            .SetStyle(new RowStyle { RowHeight = 20 })
            .Build(),

        RowBuilder
            .SetCells(
                CellBuilder.SetLocation("A", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                CellBuilder.SetLocation("B", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                CellBuilder.SetLocation("C", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .SetValue("Before Sharing").Build(),
                CellBuilder.SetLocation("D", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .SetValue("After Sharing").Build(),
                CellBuilder.SetLocation("E", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .SetValue("Sum").Build(),
                CellBuilder.SetLocation("F", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .SetValue("Before Sharing").Build(),
                CellBuilder.SetLocation("G", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .SetValue("After Sharing").Build(),
                CellBuilder.SetLocation("H", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .SetValue("Sum").Build(),
                CellBuilder.SetLocation("I", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build()
            )
            .NoMergedCells()
            .SetStyle(new RowStyle { RowHeight = 20 })
            .Build())
    .SetTableMergedCells(
        MergeBuilder
            .SetMergingStartPoint("A", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
            .SetMergingFinishPoint("A", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
            .Build(),
        MergeBuilder
            .SetMergingStartPoint("B", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
            .SetMergingFinishPoint("B", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
            .Build(),

        MergeBuilder
            .SetMergingStartPoint("F", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
            .SetMergingFinishPoint("H", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
            .SetMergingAreaBackgroundColor(Color.DarkBlue)
            .Build(),
        MergeBuilder
            .SetMergingStartPoint("I", tableCreditsDebits.GetNextVerticalRowNumberAfterTable())
            .SetMergingFinishPoint("I", tableCreditsDebits.GetNextVerticalRowNumberAfterTable() + 1)
            .SetMergingAreaBackgroundColor(Color.Yellow)
            .Build()
    )

    .SetStyle(new TableStyle
    {
        BackgroundColor = Color.Blue,
        Font = new TextFont { FontColor = Color.White }
    })
    .Build();
```

**4- Table: with Salaries data with thin borders**
```csharp
var tableSalaries = TableBuilder
            .CreateStepByStepManually()
            .SetRows(accountsReportDto.AccountSalaryCodes.Select((account, index) =>
                RowBuilder
                    .SetCells(
                        CellBuilder.SetLocation("A", tableBlueBg.GetNextVerticalRowNumberAfterTable() + index)
                            .SetValue(account.Name).Build(),
                        CellBuilder.SetLocation("B", tableBlueBg.GetNextVerticalRowNumberAfterTable() + index)
                            .SetValue(account.Code).Build()
                        )
                    .NoMergedCells()
                    .NoCustomStyle()
                    .Build()
            ).ToArray())
            .HasNoMergedCells()
            .SetStyle(new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black }
            })
            .Build();
```
**5- Table:  Sharing info**
Table with sharing before/after data
```csharp
        var tableSharingBeforeAfterData = TableBuilder
            .CreateStepByStepManually()
            .SetRows(RowBuilder
                    .SetCells(
                        CellBuilder
                            .SetLocation("C", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                            .SetValue(accountsReportDto.AccountSharingData
                                .Where(s => s.AccountName == "Branch 1")
                                .Select(s => s.AccountSharingDetail.BeforeSharing)
                                .FirstOrDefault())
                            .Build(),

                        CellBuilder
                            .SetLocation("D", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                            .SetValue(accountsReportDto.AccountSharingData
                                .Where(s => s.AccountName == "Branch 1")
                                .Select(s => s.AccountSharingDetail.AfterSharing)
                                .FirstOrDefault())
                            .Build(),

                        CellBuilder
                            .SetLocation("E", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                            .SetValue(accountsReportDto.AccountSharingData
                                .Where(s => s.AccountName == "Branch 1")
                                .Select(s => s.AccountSharingDetail.AfterSharing + s.AccountSharingDetail.BeforeSharing)
                                .FirstOrDefault())
                            .Build(),

                        CellBuilder
                            .SetLocation("F", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                            .SetValue(11000)
                            .Build(),

                        CellBuilder
                            .SetLocation("G", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                            .SetValue(10000)
                            .Build(),

                        CellBuilder
                            .SetLocation("H", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                            .SetValue(21000)
                            .Build(),

                        CellBuilder
                            .SetLocation("I", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                            .SetValue(accountsReportDto.Average)
                            .Build()
                        )
                    .NoMergedCells()
                    .NoCustomStyle()
                    .Build())
            .SetTableMergedCells(
                MergeBuilder
                    .SetMergingStartPoint("C", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("C", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("D", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("D", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("E", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("E", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("F", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("F", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("G", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("G", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("H", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("H", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("I", tableBlueBg.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("I", tableBlueBg.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build()
            )
            .NoCustomStyle()
            .Build();
```

**6- Row: Light Green row for report date**

```csharp
IRowBuilder rowReportDate = RowBuilder
    .SetCells(
        CellBuilder
            .SetLocation("D", tableSharingBeforeAfterData.GetNextVerticalRowNumberAfterTable() + 1)
            .SetValue(DateTime.Now)
            .Build()
        )
    .SetRowMergedCells(MergeBuilder
        .SetMergingStartPoint("D", tableSharingBeforeAfterData.GetNextVerticalRowNumberAfterTable() + 1)
        .SetMergingFinishPoint("F", tableSharingBeforeAfterData.GetNextVerticalRowNumberAfterTable() + 1)
        .Build())
    .NoCustomStyle()
    .Build();

```

**7- Cell: User name (me!)**

```csharp
ICellBuilder cellUserName = CellBuilder
    .SetLocation("E", rowReportDate.GetNextRowNumberAfterRow() + 1)
    .SetValue("Farshad Davoudi")
    .SetStyle(new CellStyle
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
    })
    .Build();
```

## *3- Create `IExcelBuilder`*

Then we create our main model by using the Excel sub-sections builders created in Step 2 plus other styles that are available in this builder.

```csharp
var excelBuilder = ExcelBuilder
    .SetGeneratedFileName("Accounts Report")
    .CreateComplexLayoutExcel()
    .SetSheets(SheetBuilder
        .SetName("Sheet1")
        .SetTable(tableTopHeader)
        .SetTable(tableCreditsDebits)
        .SetTable(tableBlueBg)
        .SetTable(tableSalaries)
        .SetTable(tableSharingBeforeAfterData)
        .SetRow(rowReportDate)
        .SetCell(cellUserName)
        .NoMoreTablesRowsOrCells()
        .NoCustomStyle()
        .Build())
    .SetSheetsDefaultStyle(new SheetsDefaultStyle { AllSheetsDefaultTextAlign = TextAlign.Center })
    .Build();
```

## *4- Finally Generate Excel using `ExcelWizardService`*

At last, we create our gorgeous Excel! by injecting `IExcelWizardService` and use one of its methods. It is the easiest part! 

```csharp
return Ok(_excelWizardService.GenerateExcel(excelBuilder, @"C:\GeneratedExcelSamples"));
```
