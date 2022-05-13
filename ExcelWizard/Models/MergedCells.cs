namespace ExcelWizard.Models;

public class MergedCells
{
    /// <summary>
    /// Merged Cells Start and End Location
    /// </summary>
    public MergedBoundaryLocation MergedBoundaryLocation { get; set; } = new();

    /// <summary>
    /// Set outside border of a Merged Cells (like table). Default will inherit
    /// </summary>
    public Border? OutsideBorder { get; set; }
}