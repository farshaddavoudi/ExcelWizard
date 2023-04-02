using ExcelWizard.Models.EWCell;

namespace ExcelWizard.Tests.Models.EWCell;

public class CellLocationTests
{
    [Theory]
    [InlineData(2, 1, "B1")]
    [InlineData("C", 10, "C10")]
    [InlineData(1, 5, "A5")]
    public void GetCellLocationDisplayName_GivenTheColumnAndRowInConstructor_ShouldReturnLocDisplayName(
        dynamic columnLetterOrNumber, int rowNumber, string expectedDisplay)
    {
        // Arrange 
        var cellLoc = new CellLocation(columnLetterOrNumber, rowNumber);

        // Act
        var actualDisplayName = cellLoc.GetCellLocationDisplayName();

        // Assert
        actualDisplayName.Should().Be(expectedDisplay);
    }
}