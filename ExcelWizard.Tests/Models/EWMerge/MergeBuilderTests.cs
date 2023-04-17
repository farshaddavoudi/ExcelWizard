using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace ExcelWizard.Tests.Models.EWMerge;

public class MergeBuilderTests
{
    [Fact]
    public void SetMergingStartPoint_GivenStartPointXAndY_ShouldSetCorrectlyOnMergedCellsModel()
    {
        // Arrange
        var columnLetter = "B"; var columnNo = 2;
        var rowNo = new Faker().Random.Int(1, 10);

        // Act
        var mergeBuilder = MergeBuilder
            .SetMergingStartPoint(columnLetter, rowNo)
            .SetMergingFinishPoint(It.IsAny<int>(), It.IsAny<int>())
            .Build();

        var mergeCells = (MergedCells)mergeBuilder;

        // Assert
        mergeCells.Should().NotBeNull();

        mergeCells.MergedBoundaryLocation.StartCellLocation.Should().NotBeNull();

        mergeCells.MergedBoundaryLocation.StartCellLocation!.ColumnNumber.Should().Be(columnNo);

        mergeCells.MergedBoundaryLocation.StartCellLocation!.RowNumber.Should().Be(rowNo);
    }

    [Fact]
    public void SetMergingFinishPoint_GivenFinishPointXAndY_ShouldSetCorrectlyOnMergedCellsModel()
    {
        // Arrange
        var columnLetter = "D"; var columnNo = 4;
        var rowNo = new Faker().Random.Int(5, 10);

        // Act
        var mergeBuilder = MergeBuilder
            .SetMergingStartPoint(It.IsAny<int>(), It.IsAny<int>())
            .SetMergingFinishPoint(columnLetter, rowNo)
            .Build();

        var mergeCells = (MergedCells)mergeBuilder;

        // Assert
        mergeCells.Should().NotBeNull();

        mergeCells.MergedBoundaryLocation.FinishCellLocation.Should().NotBeNull();

        mergeCells.MergedBoundaryLocation.FinishCellLocation!.ColumnNumber.Should().Be(columnNo);

        mergeCells.MergedBoundaryLocation.FinishCellLocation!.RowNumber.Should().Be(rowNo);
    }

    [Fact]
    public void SetMergingAreaBackgroundColor_GivenRedColor_ShouldSetRedForMergeCellsModel()
    {
        // Arrange
        var color = System.Drawing.Color.Red;

        // Act
        var mergeBuilder = MergeBuilder
            .SetMergingStartPoint(It.IsAny<int>(), It.IsAny<int>())
            .SetMergingFinishPoint(It.IsAny<int>(), It.IsAny<int>())
            .SetMergingAreaBackgroundColor(color)
            .Build();

        var mergeCells = (MergedCells)mergeBuilder;

        // Assert
        mergeCells.Should().NotBeNull();

        mergeCells.BackgroundColor.Should().NotBeNull();

        mergeCells.BackgroundColor.Should().Be(color);
    }

    [Fact]
    public void SetMergingOutsideBorder_GivenNoBorderColor_ShouldSetBlackForBorderColor()
    {
        // Act
        var mergeBuilder = MergeBuilder
            .SetMergingStartPoint(It.IsAny<int>(), It.IsAny<int>())
            .SetMergingFinishPoint(It.IsAny<int>(), It.IsAny<int>())
            .SetMergingOutsideBorder(new Faker().PickRandom<LineStyle>())
            .Build();

        var mergeCells = (MergedCells)mergeBuilder;

        // Assert
        mergeCells.Should().NotBeNull();

        mergeCells.OutsideBorder.Should().NotBeNull();

        mergeCells.OutsideBorder!.BorderColor.Should().Be(Color.Black);
    }

    [Fact]
    public void SetMergingOutsideBorder_GivenLineStyleAndRedColor_ShouldSetCorrectlyOnMergeCellsModel()
    {
        // Arrange
        var borderLineStyle = new Faker().PickRandom<LineStyle>();
        var borderColor = Color.FromName(new Faker().Commerce.Color());

        // Act
        var mergeBuilder = MergeBuilder
            .SetMergingStartPoint(It.IsAny<int>(), It.IsAny<int>())
            .SetMergingFinishPoint(It.IsAny<int>(), It.IsAny<int>())
            .SetMergingOutsideBorder(borderLineStyle, borderColor)
            .Build();

        var mergeCells = (MergedCells)mergeBuilder;

        // Assert
        mergeCells.Should().NotBeNull();

        mergeCells.OutsideBorder.Should().NotBeNull();

        mergeCells.OutsideBorder!.BorderColor.Should().Be(borderColor);

        mergeCells.OutsideBorder!.BorderLineStyle.Should().Be(borderLineStyle);
    }
}