using ExcelWizard.Models.EWCell;
using System.ComponentModel.DataAnnotations;

namespace ExcelWizard.Models.EWMerge;

public class MergedBoundaryLocation
{
    [Required]
    public CellLocation? StartCellLocation { get; set; }

    [Required]
    public CellLocation? FinishCellLocation { get; set; }
}