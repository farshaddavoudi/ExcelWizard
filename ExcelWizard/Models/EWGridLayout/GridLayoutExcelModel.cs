using System.Collections.Generic;

namespace ExcelWizard.Models.EWGridLayout;

public record GridLayoutExcelModel(string? GeneratedFileName, List<BoundSheet> Sheets);