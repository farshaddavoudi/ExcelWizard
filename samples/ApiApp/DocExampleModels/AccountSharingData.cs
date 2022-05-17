namespace ApiApp.DocExampleModels;

public class AccountSharingData
{
    public string? AccountName { get; set; }

    public List<AccountSharingDetail> AccountSharingDetailsList { get; set; } = new();
}

public class AccountSharingDetail
{
    public decimal BeforeSharing { get; set; }

    public decimal AfterSharing { get; set; }
}