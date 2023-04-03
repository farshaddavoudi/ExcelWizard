namespace ExcelWizard.Tests.Models.EWCell;

public class CellStyleTests : IClassFixture<CellStyleFixture>
{
    private readonly CellStyleFixture _fixture;

    public CellStyleTests(CellStyleFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void NewInstance_FontShouldBeNull()
    {
        // Assert
        _fixture.CellStyle.Font.Should().BeNull();
    }

    [Fact]
    public void NewInstance_WordwrapShouldBeFalse()
    {
        // Assert
        _fixture.CellStyle.Wordwrap.Should().BeFalse();
    }

    [Fact]
    public void NewInstance_CellTextAlignShouldBeNull()
    {
        // Assert
        _fixture.CellStyle.CellTextAlign.Should().BeNull();
    }

    [Fact]
    public void NewInstance_BackgroundColorShouldBeNull()
    {
        // Assert
        _fixture.CellStyle.BackgroundColor.Should().BeNull();
    }

    [Fact]
    public void NewInstance_CellBorderShouldBeNull()
    {
        // Assert
        _fixture.CellStyle.CellBorder.Should().BeNull();
    }
}