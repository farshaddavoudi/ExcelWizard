using System.Drawing;

namespace ExcelWizard.Models.EWStyles;

public class Border
{
    #region Ctors

    public Border() { }

    public Border(LineStyle borderLineStyle, Color borderColor)
    {
        BorderLineStyle = borderLineStyle;

        BorderColor = borderColor;
    }

    #endregion

    public LineStyle BorderLineStyle { get; set; } = LineStyle.None;

    public Color BorderColor { get; set; } = Color.Black;
}