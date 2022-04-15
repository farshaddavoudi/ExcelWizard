namespace EasyExcelGenerator.Models;

public class GeneratedExcelFile
{
    public byte[]? Content { get; set; }

    public const string MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public const string Extension = "xlsx";
}