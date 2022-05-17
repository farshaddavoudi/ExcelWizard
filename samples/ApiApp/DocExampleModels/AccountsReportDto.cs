namespace ApiApp.DocExampleModels;

public class AccountsReportDto
{
    public string? ReportName { get; set; }

    public List<AccountDebitCredit> AccountDebitCreditList { get; set; } = new();

    public List<AccountSalaryCode> AccountSalaryCodes { get; set; } = new();

    public List<AccountSharingData> AccountSharingData { get; set; } = new();
}