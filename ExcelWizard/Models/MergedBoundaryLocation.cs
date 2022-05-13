using System.ComponentModel.DataAnnotations;

namespace ExcelWizard.Models;

public class MergedBoundaryLocation
{
    [Required]
    public CellLocation? FirstCellLocation { get; set; }

    [Required]
    public CellLocation? LastCellLocation { get; set; }
}