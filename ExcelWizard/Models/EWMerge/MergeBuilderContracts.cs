using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace ExcelWizard.Models.EWMerge;

public interface IMergeBuilder
{

}

public interface IExpectMergingFinishPointMergeBuilder
{
    IExpectStylesOrBuildMergeBuilder SetMergingFinishPoint(dynamic columnLetterOrNumber, int rowNumber);
}

public interface IExpectStylesOrBuildMergeBuilder
{
    IExpectStylesOrBuildMergeBuilder SetMergingAreaBackgroundColor(Color backgroundColor);

    IExpectStylesOrBuildMergeBuilder SetMergingOutsideBorderStyle(LineStyle borderLineStyle = LineStyle.Thin, Color borderColor = new());

    MergedCells Build();
}