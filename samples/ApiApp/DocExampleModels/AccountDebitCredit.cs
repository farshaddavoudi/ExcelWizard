using ExcelWizard.Models;
using ExcelWizard.Models.EWStyles;
using ExcelWizard.Models.EWTable;
using System.Drawing;

namespace ApiApp.DocExampleModels;

[ExcelTable(HeaderBackgroundColor = KnownColor.LightGray,
    //HeaderOccupyingRowsNo = 3,
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


