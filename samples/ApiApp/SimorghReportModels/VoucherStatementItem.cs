namespace ApiApp.SimorghReportModels;

public class VoucherStatementItem
{
    public string? AccountCode { get; set; }

    public decimal Debit { get; set; }

    public decimal Credit { get; set; }
}