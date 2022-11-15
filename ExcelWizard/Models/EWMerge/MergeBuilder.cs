using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWStyles;
using System;
using System.Drawing;

namespace ExcelWizard.Models.EWMerge;

public class MergeBuilder : IExpectMergingFinishPointMergeBuilder, IExpectStylesOrBuildMergeBuilder
{
    private MergeBuilder() { }

    private MergedCells MergedCells { get; set; } = new();
    private bool CanBuild { get; set; }

    /// <summary>
    /// Set merging start location cell
    /// </summary>
    /// <param name="columnLetterOrNumber"> Start location column letter or number, e.g. "A" or 1 </param>
    /// <param name="rowNumber"> Start location row number, e.g. 1 </param>
    public static IExpectMergingFinishPointMergeBuilder SetMergingStartPoint(dynamic columnLetterOrNumber, int rowNumber)
    {
        return new MergeBuilder
        {
            MergedCells = new MergedCells
            {
                MergedBoundaryLocation = new MergedBoundaryLocation
                {
                    StartCellLocation = new CellLocation(columnLetterOrNumber, rowNumber)
                }
            }
        };
    }

    public IExpectStylesOrBuildMergeBuilder SetMergingFinishPoint(dynamic columnLetterOrNumber, int rowNumber)
    {
        CanBuild = true;

        MergedCells.MergedBoundaryLocation.FinishCellLocation = new CellLocation(columnLetterOrNumber, rowNumber);

        return this;
    }

    public IExpectStylesOrBuildMergeBuilder SetMergingAreaBackgroundColor(Color backgroundColor)
    {
        MergedCells.BackgroundColor = backgroundColor;

        return this;
    }

    public IExpectStylesOrBuildMergeBuilder SetMergingOutsideBorder(LineStyle borderLineStyle = LineStyle.Thin, Color borderColor = new())
    {
        if (borderColor.IsEmpty)
            borderColor = Color.Black;

        MergedCells.OutsideBorder = new Border(borderLineStyle, borderColor);

        return this;
    }

    public MergedCells Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build MergedCells because some necessary information are not provided");

        return MergedCells;
    }
}