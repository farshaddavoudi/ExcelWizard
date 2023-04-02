using ExcelWizard.Models;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWStyles;
using System.Drawing;

namespace ExcelWizard.Tests.Models.EWCell;

public class CellBuilderTests
{
    [Theory]
    [InlineData(1, 1, 2)]
    [InlineData("B", 2, 3)]
    public void SetLocation_WhenGivenColumnAndRow_ShouldSetCellLocation(dynamic columnLetterOrNumber, int columnNumber, int rowNumber)
    {
        // Act
        Cell cell = CellBuilder
            .SetLocation(columnLetterOrNumber, rowNumber)
            .Build();

        // Assert
        cell.CellLocation.RowNumber.Should().Be(rowNumber);
        cell.CellLocation.ColumnNumber.Should().Be(columnNumber);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(1)]
    [InlineData("foo")]
    public void SetValue_WhenGivenAValue_ShouldSetItAsCellValue(object? value)
    {
        // Act
        Cell cell = CellBuilder
            .SetLocation(It.IsAny<int>(), It.IsAny<int>())
            .SetValue(value)
            .Build();

        // Assert
        cell.CellValue.Should().BeEquivalentTo(value);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("foo")]
    public void SetIdentifier_WhenGivenAnIdentifier_ShouldSetItAsCellIdentifier(string? identifier)
    {
        // Act
        Cell cell = CellBuilder
            .SetLocation(It.IsAny<int>(), It.IsAny<int>())
            .SetIdentifier(identifier)
            .Build();

        // Assert
        cell.CellIdentifier.Should().Be(identifier);
    }

    [Fact]
    public void SetContentType_WhenNotSpecifyAnyType_ShouldSetDefaultGeneralType()
    {
        // Act
        Cell cell = CellBuilder
            .SetLocation(It.IsAny<int>(), It.IsAny<int>())
            .Build();

        // Assert
        cell.CellContentType.Should().Be(CellContentType.General);
    }

    [Fact]
    public void SetContentType_WhenGivenAType_ShouldSetItAsCellContentType()
    {
        // Arrange
        var contentType = CellContentType.Number;

        // Act
        Cell cell = CellBuilder
            .SetLocation(It.IsAny<int>(), It.IsAny<int>())
            .SetContentType(contentType)
            .Build();

        // Assert
        cell.CellContentType.Should().Be(contentType);
    }

    [Fact]
    public void SetCellStyle_WhenGivenStyle_ShouldSetItAsCellStyle()
    {
        // Arrange
        var cellStyle = new CellStyle { Wordwrap = true, BackgroundColor = Color.Red, Font = new TextFont { FontName = "Foo" } };

        // Act
        Cell cell = CellBuilder
            .SetLocation(It.IsAny<int>(), It.IsAny<int>())
            .SetCellStyle(cellStyle)
            .Build();

        // Assert
        cell.CellStyle.Should().Be(cellStyle);
    }

    [Fact]
    public void SetCellContentVisibility_WhenGivenIsVisibleFalse_ShouldSetCellIsContentVisibleTrue()
    {
        // Act
        Cell cell = CellBuilder
            .SetLocation(It.IsAny<int>(), It.IsAny<int>())
            .Build();

        // Assert
        cell.IsCellVisible.Should().BeTrue();
    }

    [Fact]
    public void SetCellContentVisibility_WhenContentVisibilityIsNotSet_ShouldSetIsContentVisible()
    {
        var isVisible = false;

        // Act
        Cell cell = CellBuilder
            .SetLocation(It.IsAny<int>(), It.IsAny<int>())
            .SetContentVisibility(isVisible)
            .Build();

        // Assert
        cell.IsCellVisible.Should().BeFalse();
    }
}