using ExcelWizard.Models;
using ExcelWizard.Models.EWStyles;
using ExcelWizard.Models.EWTable;
using System.Drawing;

namespace ApiApp.DocExampleModels;

[ExcelTable(HeaderBackgroundColor = KnownColor.LightGray,
    InsideCellsBorderStyle = LineStyle.Thick,
    InsideCellsBorderColor = KnownColor.Black,
    OutsideBorderColor = KnownColor.Black,
    OutsideBorderStyle = LineStyle.Thick,
    TextAlign = TextAlign.Center)]
public class AccountDebitCredit
{
    [ExcelTableColumn(HeaderName = "Account Code")]
    public string? AccountCode { get; set; }

    [ExcelTableColumn(DataContentType = CellContentType.Currency)]
    public decimal Debit { get; set; }

    [ExcelTableColumn(DataContentType = CellContentType.Currency)]
    public decimal Credit { get; set; }
}