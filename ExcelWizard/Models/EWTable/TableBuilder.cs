using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWStyles;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace ExcelWizard.Models.EWTable;

public class TableBuilder : IExpectRowsTableBuilder, IExpectMergedCellsStatusInManualProcessTableBuilder,
    IExpectStyleTableBuilder, IExpectMergedCellsStatusInModelTableBuilder,
    IExpectBuildMethodInModelTableBuilder, IExpectBuildMethodInManualTableBuilder
{
    private TableBuilder() { }

    private Table Table { get; set; } = new();
    private bool CanBuild { get; set; }

    /// <summary>
    /// Automatically construct the Table using a model data and attributes. Attributes to configure are [ExcelTable] and [ExcelTableColumn]
    /// </summary>
    /// <param name="bindingListModel">The model instance which should be list of an item. The type should be configured by attributes for some styles and other configs </param>
    /// <param name="tableStartPoint"> The start location of the table. The end point will be calculated dynamically </param>
    public static IExpectMergedCellsStatusInModelTableBuilder CreateUsingAModelToBind(object bindingListModel, CellLocation tableStartPoint)
    {
        var isObjectDataList = bindingListModel is IEnumerable;

        if (isObjectDataList is false)
            throw new InvalidOperationException("Provided data for table is not a valid data list");

        var headerRow = new Row();

        var dataRows = new List<Row>();

        // Get Header 

        bool isHeaderAlreadyCalculated = false;

        bool hasHeader = true;

        List<MergedCells> tableHeaderMerges = new();

        int yLocation = tableStartPoint.RowNumber;

        var borderType = LineStyle.Thin;

        Border outsideBorder = new();

        Border insideBorder = new();

        if (bindingListModel is IEnumerable records)
        {
            foreach (var record in records)
            {
                // Each record is an entire row of Excel

                var excelTableAttribute = record.GetType().GetCustomAttribute<ExcelTableAttribute>();

                hasHeader = excelTableAttribute?.HasHeader ?? true;

                var headerOccupyingRowsNo = excelTableAttribute?.HeaderOccupyingRowsNo ?? 1;

                var tableDefaultFontWeight = excelTableAttribute?.FontWeight;

                var tableDefaultFont = new TextFont
                {
                    FontName = excelTableAttribute?.FontName,
                    FontSize = excelTableAttribute?.FontSize == 0 ? null : excelTableAttribute?.FontSize,
                    FontColor = Color.FromKnownColor(excelTableAttribute?.FontColor ?? KnownColor.Black),
                    IsBold = tableDefaultFontWeight == FontWeight.Bold
                };

                outsideBorder = new Border(excelTableAttribute?.OutsideBorderStyle ?? LineStyle.Thin,
                    Color.FromKnownColor(excelTableAttribute?.OutsideBorderColor ?? KnownColor.LightGray));

                insideBorder = new Border(excelTableAttribute?.InsideCellsBorderStyle ?? LineStyle.Thin,
                    Color.FromKnownColor(excelTableAttribute?.InsideCellsBorderColor ?? KnownColor.LightGray));

                TextAlign tableDefaultTextAlign = excelTableAttribute?.TextAlign ?? TextAlign.Inherit;

                PropertyInfo[] properties = record.GetType().GetProperties();

                int xLocation = tableStartPoint.ColumnNumber;

                var recordRow = new Row
                {
                    RowStyle = new RowStyle
                    {
                        BackgroundColor = Color.FromKnownColor(excelTableAttribute?.DataBackgroundColor ?? KnownColor.Transparent)
                    }
                };

                // Each loop is a Column

                foreach (var prop in properties)
                {
                    var excelTableColumnAttribute = (ExcelTableColumnAttribute?)prop.GetCustomAttributes(true).FirstOrDefault(x => x is ExcelTableColumnAttribute);

                    if (excelTableColumnAttribute?.Ignore ?? false)
                        continue;

                    TextAlign GetCellTextAlign(TextAlign defaultAlign, TextAlign? headerOrDataTextAlign)
                    {
                        return headerOrDataTextAlign switch
                        {
                            TextAlign.Inherit => defaultAlign,
                            _ => headerOrDataTextAlign ?? defaultAlign
                        };
                    }

                    var finalFont = new TextFont
                    {
                        FontName = excelTableColumnAttribute?.FontName ?? tableDefaultFont.FontName,
                        FontSize = excelTableColumnAttribute?.FontSize is null || excelTableColumnAttribute.FontSize == 0 ? tableDefaultFont.FontSize : excelTableColumnAttribute.FontSize,
                        FontColor = excelTableColumnAttribute is null || excelTableColumnAttribute.FontColor == KnownColor.Transparent
                            ? tableDefaultFont.FontColor.Value
                            : Color.FromKnownColor(excelTableColumnAttribute.FontColor),
                        IsBold = excelTableColumnAttribute is null || excelTableColumnAttribute.FontWeight == FontWeight.Inherit
                            ? tableDefaultFont.IsBold
                            : excelTableColumnAttribute.FontWeight == FontWeight.Bold
                    };

                    // Header
                    if (hasHeader && isHeaderAlreadyCalculated is false)
                    {
                        var isBold = excelTableColumnAttribute is null ||
                                     excelTableColumnAttribute.FontWeight == FontWeight.Inherit
                            ? tableDefaultFontWeight != FontWeight.Normal
                            : excelTableColumnAttribute.FontWeight == FontWeight.Bold;

                        var headerFontColor = excelTableAttribute?.HeaderFontColor != null ? Color.FromKnownColor(excelTableAttribute.HeaderFontColor) : Color.Empty;

                        var headerFontWeight = excelTableAttribute?.HeaderFontWeight != null ? excelTableAttribute.FontWeight : FontWeight.Inherit;

                        var headerFont = new TextFont
                        {
                            FontColor = headerFontColor == Color.Empty ? finalFont.FontColor : headerFontColor,
                            FontName = finalFont.FontName,
                            FontSize = finalFont.FontSize,
                            IsBold = headerFontWeight == FontWeight.Inherit ? isBold : headerFontWeight == FontWeight.Bold
                        };

                        Cell headerCell = CellBuilder
                            .SetLocation(xLocation, yLocation)
                            .SetValue(excelTableColumnAttribute?.HeaderName ?? prop.Name)
                            .SetCellStyle(new CellStyle
                            {
                                Font = headerFont,
                                CellTextAlign = GetCellTextAlign(tableDefaultTextAlign,
                                    excelTableColumnAttribute?.HeaderTextAlign)
                            })
                            .SetContentType(CellContentType.Text)
                            .Build();

                        headerRow.RowCells.Add(headerCell);

                        var headerBgColor = excelTableAttribute?.HeaderBackgroundColor != null ? Color.FromKnownColor(excelTableAttribute.HeaderBackgroundColor) : Color.Transparent;

                        headerRow.RowStyle.BackgroundColor = headerBgColor;

                        headerRow.RowStyle.RowOutsideBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                        headerRow.RowStyle.InsideCellsBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                        if (headerOccupyingRowsNo > 1)
                        {
                            tableHeaderMerges.Add(new MergedCells
                            {
                                BackgroundColor = headerBgColor,

                                MergedBoundaryLocation = new MergedBoundaryLocation
                                {
                                    StartCellLocation = new CellLocation(xLocation, tableStartPoint.RowNumber),
                                    FinishCellLocation = new CellLocation(xLocation, tableStartPoint.RowNumber + headerOccupyingRowsNo - 1)
                                }
                            });
                        }
                    }

                    // Data
                    int dataYLocation = hasHeader ? yLocation + headerOccupyingRowsNo : yLocation;

                    var dataCell = CellBuilder
                        .SetLocation(xLocation, dataYLocation)
                        .SetValue(prop.GetValue(record))
                        .SetContentType(excelTableColumnAttribute?.DataContentType ?? CellContentType.Text)
                        .SetCellStyle(new CellStyle
                        {
                            Font = finalFont,
                            CellTextAlign = GetCellTextAlign(tableDefaultTextAlign,
                                excelTableColumnAttribute?.DataTextAlign)
                        })
                        .Build();

                    recordRow.RowCells.Add(dataCell);

                    xLocation++;
                }

                dataRows.Add(recordRow);

                yLocation++;

                isHeaderAlreadyCalculated = true;
            }
        }

        // End of calculations 

        List<Row> allRows = new List<Row>();

        if (hasHeader)
            allRows.Add(headerRow);

        allRows.AddRange(dataRows);

        return new TableBuilder
        {
            Table = new Table
            {
                TableRows = allRows,
                TableStyle = new TableStyle
                {
                    TableOutsideBorder = outsideBorder,
                    InsideCellsBorder = insideBorder
                },
                MergedCellsList = tableHeaderMerges
            }
        };
    }

    /// <summary>
    /// Manually build the Table defining its properties and styles step by step
    /// </summary>
    public static IExpectRowsTableBuilder CreateStepByStepManually()
    {
        return new TableBuilder
        {
            Table = new Table()
        };
    }

    public IExpectMergedCellsStatusInManualProcessTableBuilder SetRows(IRowBuilder rowBuilder, params IRowBuilder[] rowBuilders)
    {
        IRowBuilder[] rows = new[] { rowBuilder }.Concat(rowBuilders).ToArray();

        Table.TableRows = rows.Select(r => (Row)r).ToList();

        return this;
    }

    public IExpectMergedCellsStatusInManualProcessTableBuilder SetRows(List<IRowBuilder> rowBuilders)
    {
        Table.TableRows = rowBuilders.Select(r => (Row)r).ToList();

        return this;
    }

    public IExpectStyleTableBuilder SetTableMergedCells(IMergeBuilder mergeBuilder, params IMergeBuilder[] mergeBuilders)
    {
        IMergeBuilder[] merges = new[] { mergeBuilder }.Concat(mergeBuilders).ToArray();

        CanBuild = true;

        Table.MergedCellsList = merges.Select(m => (MergedCells)m).ToList();

        return this;
    }

    public IExpectStyleTableBuilder SetTableMergedCells(List<IMergeBuilder> mergeBuilders)
    {
        CanBuild = true;

        Table.MergedCellsList = mergeBuilders.Select(m => (MergedCells)m).ToList();

        return this;
    }

    public IExpectStyleTableBuilder TableHasNoMerging()
    {
        CanBuild = true;

        return this;
    }

    public IExpectBuildMethodInManualTableBuilder SetTableStyle(TableStyle tableStyle)
    {
        Table.TableStyle = tableStyle;

        return this;
    }

    public IExpectBuildMethodInManualTableBuilder TableHasNoCustomStyle()
    {
        return this;
    }

    public IExpectBuildMethodInModelTableBuilder SetBoundTableMergedCells(IMergeBuilder mergeBuilder, params IMergeBuilder[] mergeBuilders)
    {
        IMergeBuilder[] merges = new[] { mergeBuilder }.Concat(mergeBuilders).ToArray();

        CanBuild = true;

        Table.MergedCellsList = merges.Select(m => (MergedCells)m).ToList();

        return this;
    }

    public IExpectBuildMethodInModelTableBuilder SetBoundTableMergedCells(List<IMergeBuilder> mergeBuilders)
    {
        CanBuild = true;

        Table.MergedCellsList = mergeBuilders.Select(m => (MergedCells)m).ToList();

        return this;
    }

    public IExpectBuildMethodInModelTableBuilder BoundTableHasNoMerging()
    {
        CanBuild = true;

        return this;
    }

    public ITableBuilder Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Table because some necessary information not provided");

        return Table;
    }
}