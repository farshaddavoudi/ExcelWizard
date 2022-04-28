namespace ExcelWizard.Models;

public class ProtectionLevel
{
    public string? Password { get; set; }

    public bool SelectLockedCells { get; set; } = true;

    public bool SelectUnlockedCells { get; set; } = true;

    public bool FormatCells { get; set; } = true;

    public bool FormatColumns { get; set; } = true;

    public bool FormatRows { get; set; } = true;

    public bool InsertColumns { get; set; } = true;

    public bool InsertRows { get; set; } = true;

    public bool InsertHyperLinks { get; set; } = true;

    public bool DeleteColumns { get; set; } = true;

    public bool DeleteRows { get; set; } = true;

    public bool Sort { get; set; } = true;

    public bool UseAutoFilter { get; set; } = true;

    public bool UsePivotTableReports { get; set; } = true;

    public bool EditObjects { get; set; } = true;

    public bool EditScenarios { get; set; } = true;

    public bool HardProtect { get; set; } = false; // Will disable all the elements and hardly protect the Sheet
}