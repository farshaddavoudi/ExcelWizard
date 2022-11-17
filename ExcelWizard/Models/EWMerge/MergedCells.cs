using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace ExcelWizard.Models.EWMerge;

public class MergedCells : IMergeBuilder
{
    /// <summary>
    /// Merged Cells Start and End Location
    /// </summary>
    public MergedBoundaryLocation MergedBoundaryLocation { get; internal set; } = new();

    /// <summary>
    /// Set outside border of a Merged Cells (like table). Default will inherit
    /// </summary>
    public Border? OutsideBorder { get; internal set; }

    /// <summary>
    /// Set Background Color for entire Merged Cells. Default inherit
    /// </summary>
    public Color? BackgroundColor { get; internal set; }

}