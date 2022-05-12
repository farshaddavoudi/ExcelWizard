namespace ApiApp.SimorghReportModels;

public class VoucherStatementResult
{
    public string? ReportName { get; set; }

    public decimal FinalRemain { get; set; }

    public List<VoucherStatementItem> VoucherStatementItem { get; set; } = new();

    public List<AccountDto> Accounts { get; set; } = new();

    public List<SummaryAccount> SummaryAccounts { get; set; } = new();
}