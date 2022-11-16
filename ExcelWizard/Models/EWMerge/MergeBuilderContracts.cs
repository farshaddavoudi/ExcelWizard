using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace ExcelWizard.Models.EWMerge;

public interface IMergeBuilder
{

}

public interface IExpectMergingFinishPointMergeBuilder
{
    /// <summary>
    /// Set merging finish location cell
    /// </summary>
    /// <param name="columnLetterOrNumber"> Finish location column letter or number, e.g. "A" or 1 </param>
    /// <param name="rowNumber"> Finish location row number, e.g. 1 </param>
    IExpectStylesOrBuildMergeBuilder SetMergingFinishPoint(dynamic columnLetterOrNumber, int rowNumber);
}

public interface IExpectStylesOrBuildMergeBuilder
{
    /// <summary>
    /// Set Background Color for entire Merged Cells. Default inherit
    /// </summary>
    /// <param name="backgroundColor"> Merge area background color </param>
    /// <returns></returns>
    IExpectStylesOrBuildMergeBuilder SetMergingAreaBackgroundColor(Color backgroundColor);

    /// <summary>
    /// Set outside border of a Merged Cells (like table). Default will inherit
    /// </summary>
    /// <param name="borderLineStyle"></param>
    /// <param name="borderColor"></param>
    IExpectStylesOrBuildMergeBuilder SetMergingOutsideBorder(LineStyle borderLineStyle = LineStyle.Thin, Color borderColor = new());

    IMergeBuilder Build();
}