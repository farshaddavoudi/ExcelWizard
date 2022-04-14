namespace EasyExcelGenerator.Models;

public class ProtectionLevels
{
    public string? Password { get; set; }

    public bool SelectLockedCells { get; set; }

    public bool SelectUnlockedCells { get; set; }

    public bool FormatCells { get; set; }

    public bool FormatColumns { get; set; }

    public bool FormatRows { get; set; }

    public bool InsertColumns { get; set; }

    public bool InsertRows { get; set; }

    public bool InsertHyperLinks { get; set; }

    public bool DeleteColumns { get; set; }

    public bool DeleteRows { get; set; }

    public bool Sort { get; set; }

    public bool UseAutoFilter { get; set; }

    public bool UsePivotTableReports { get; set; }

    public bool EditObjects { get; set; }

    public bool EditScenarios { get; set; }
}