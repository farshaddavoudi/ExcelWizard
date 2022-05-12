namespace ApiApp.SimorghReportModels;

public class SummaryAccount
{
    public string? AccountName { get; set; }

    public List<Multiplex> Multiplex { get; set; } = new();
}