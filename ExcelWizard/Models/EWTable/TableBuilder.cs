using ExcelWizard.Models.EWRow;
using System;
using System.Collections.Generic;

namespace ExcelWizard.Models.EWTable;

public class TableBuilder : ITableBuilder, IExpectRowsTableBuilder, IExpectMergedCellsStatusInManualProcessTableBuilder,
    IExpectStyleTableBuilder, IExpectMergedCellsStatusInModelTableBuilder
{
    private TableBuilder() { }

    private Table Table { get; set; } = new();
    private bool CanBuild { get; set; }

    /// <summary>
    /// Automatically build the Table using the model data and attributes. Model should be an IEnumerable object i.e. list of an item
    /// </summary>
    /// <param name="dataList">The model instance of list of an item. The type should b configured by attributes for some styles and other configs </param>
    /// <param name="tableStartPoint"> The start location of the table. The end point will be calculated dynamically </param>
    public static IExpectMergedCellsStatusInModelTableBuilder ConstructUsingModelAutomatically(object dataList, CellLocation tableStartPoint)
    {
        return new TableBuilder
        {
            Table = new Table()
        };
    }

    /// <summary>
    /// Manually build the Table defining its properties and styles step by step
    /// </summary>
    public static IExpectRowsTableBuilder ConstructStepByStepManually()
    {
        return new TableBuilder
        {
            Table = new Table()
        };
    }

    public IExpectMergedCellsStatusInManualProcessTableBuilder SetRows(List<Row> tableRows)
    {
        if (tableRows.Count == 0)
            throw new ArgumentException("Table instance Rows cannot be an empty list");

        Table.TableRows = tableRows;

        return this;
    }

    public IExpectStyleTableBuilder SetTableMergedCells(List<MergedCells> mergedCellsList)
    {
        if (mergedCellsList.Count > 0)
            CanBuild = true;

        Table.MergedCellsList = mergedCellsList;

        return this;
    }

    public IExpectStyleTableBuilder HasNoMergedCells()
    {
        CanBuild = true;

        return this;
    }

    public TableBuilder SetStyle(TableStyle tableStyle)
    {
        Table.TableStyle = tableStyle;

        return this;
    }

    public TableBuilder NoCustomStyle()
    {
        return this;
    }

    public TableBuilder SetMergedCells(List<MergedCells> mergedCellsList)
    {
        if (mergedCellsList.Count > 0)
            CanBuild = true;

        Table.MergedCellsList = mergedCellsList;

        return this;
    }

    public TableBuilder NoMergedCells()
    {
        CanBuild = true;

        return this;
    }

    public Table Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Table because some necessary information not provided");

        return Table;
    }
}