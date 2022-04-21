namespace EasyExcelGenerator.Models;

public class GeneratedExcelFile
{
    public string? FileName { get; set; }

    public byte[]? Content { get; set; }

    public string MimeType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public string Extension => "xlsx";
}