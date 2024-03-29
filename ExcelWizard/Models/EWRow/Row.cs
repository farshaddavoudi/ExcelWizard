﻿using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelWizard.Models.EWRow;

public class Row : IRowBuilder
{
    // Props

    /// <summary>
    /// Each Row contains one or more Cell(s). It is required as Row definition cannot be without Cells.
    /// </summary>
    public List<Cell> RowCells { get; internal set; } = new();

    /// <summary>
    /// Set Row Styles including Bg, Font, Height, Borders and etc
    /// </summary>
    public RowStyle RowStyle { get; internal set; } = new();

    /// <summary>
    /// Arbitrary property to define Location of Merged Cells in the current Row. The property is collection, in case
    /// we have multiple merged-cells definitions in different locations of the Row. Notice that the Merged-Cells
    /// RowNumber should match with the Row RowNumber itself, otherwise an error will throw.
    /// </summary>
    public List<MergedCells> MergedCellsList { get; internal set; } = new();

    // Methods

    /// <summary>
    /// Get Current Row Y Location (RowNumber)
    /// </summary>
    /// <returns></returns>
    public int GetRowNumber()
    {
        return RowCells.First().CellLocation.RowNumber;
    }

    public int GetNextRowNumberAfterRow()
    {
        return GetRowNumber() + 1;
    }

    /// <summary>
    /// Get current Row Starting Cell Automatically by its Cells Location
    /// </summary>
    /// <returns></returns>
    public CellLocation GetRowFirstCellLocation()
    {
        var rowColumns = RowCells.Select(c => c.CellLocation.ColumnNumber).ToList();

        var rowNumber = RowCells.First().CellLocation.RowNumber; //All Cells in a Row have equal RowNumber

        return new CellLocation(rowColumns.Min(), rowNumber);
    }

    /// <summary>
    /// Get current Row Ending Cell Automatically by its Cells Location
    /// </summary>
    /// <returns></returns>
    public CellLocation GetRowLastCellLocation()
    {
        var rowColumns = RowCells.Select(c => c.CellLocation.ColumnNumber).ToList();

        var rowNumber = RowCells.First().CellLocation.RowNumber; //All Cells in a Row have equal RowNumber

        return new CellLocation(rowColumns.Max(), rowNumber);
    }

    public CellLocation GetNextHorizontalCellLocationAfterRow()
    {
        var rowEndLocation = GetRowLastCellLocation();

        return new CellLocation(rowEndLocation.ColumnNumber + 1, rowEndLocation.RowNumber);
    }

    public Cell? GetRowCellByColumnNumber(int columnNumber)
    {
        return RowCells.FirstOrDefault(c => c.CellLocation.ColumnNumber == columnNumber);
    }

    // Validations
    public void ValidateRowInstance()
    {
        // Row definition cannot have no Cells
        if (RowCells.Count == 0)
            throw new ValidationException("Row instance should contain one or more Cell(s)");

        // Check Y of StartLocation and EndLocation should be the equal and same with other Cells location Y property (Check with Shahab)
        if (RowCells.Select(c => c.CellLocation.RowNumber).Distinct().ToList().Count != 1)
            throw new ValidationException("Invalid Row. All Row Cells should have equal RowNumber!");

        // Check MergedCells format
        var currentRowNumber = RowCells.First().CellLocation.RowNumber;

        foreach (var cellsToMerge in MergedCellsList)
        {
            if (cellsToMerge.MergedBoundaryLocation.StartCellLocation is null || cellsToMerge.MergedBoundaryLocation.FinishCellLocation is null)
                throw new ValidationException("Something is not right about MergedCells format. FirstCellLocation and LastCellLocations cannot be null");

            if (cellsToMerge.MergedBoundaryLocation.StartCellLocation!.RowNumber != currentRowNumber)
                throw new ValidationException("Something is not right about MergedCells format. The RowNumber of FirstCellLocation should be equal with the Row RowNumber!");

            if (cellsToMerge.MergedBoundaryLocation.FinishCellLocation!.RowNumber != currentRowNumber)
                throw new ValidationException("Something is not right about MergedCells format. The RowNumber of LastCellLocation should be equal with the Row RowNumber!");
        }

        // Check all Cells be Unique (not repetitive)
        var isAllCellsUnique = RowCells.Select(c => c.CellLocation.ColumnNumber).Distinct().Count() == RowCells.Count;

        if (isAllCellsUnique is false)
            throw new ValidationException("There are some repetitive cells in the Row. All cells should be unique in a Row");

    }
}