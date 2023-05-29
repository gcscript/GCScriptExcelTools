using ClosedXML.Excel;
using GCScriptExcelTools.Models;
using System.Reflection;

namespace GCScriptExcelTools;

public static class ExcelFunctions
{
    /// <summary>
    /// Sort the worksheets in the file.
    /// <param name="workbook">The workbook to be processed.</param>
    /// <param name="ascending">True to sort ascending, false to sort descending.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool SortWorksheets(XLWorkbook workbook, bool ascending)
    {
        try
        {
            List<IXLWorksheet>? worksheets = ascending ? workbook.Worksheets.OrderBy(x => x.Name).ToList() : workbook.Worksheets.OrderByDescending(x => x.Name).ToList();

            int position = 1;
            foreach (var worksheet in worksheets)
            {
                worksheet.Position = position;
                position++;
            }

            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Remove all empty worksheets from the file.
    /// <param name="workbook">The workbook to be processed.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool RemoveEmptyWorksheets(XLWorkbook workbook)
    {
        try
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                if (worksheet.IsEmpty()) { worksheet.Delete(); }
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Remove all empty rows from the worksheet.
    /// <param name="worksheet">The worksheet to be processed.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool RemoveEmptyRows(IXLWorksheet worksheet)
    {
        try
        {
            var lastRowUsed = worksheet.LastRowUsed().RowNumber();
            for (int i = lastRowUsed; i > 0; i--)
            {
                var linha = worksheet.Row(i);
                if (linha.IsEmpty())
                {
                    linha.Delete();
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Remove all empty columns from the worksheet.
    /// <param name="worksheet">The worksheet to be processed.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool RemoveEmptyColumns(IXLWorksheet worksheet)
    {
        try
        {
            var lastColumnUsed = worksheet.LastColumnUsed().ColumnNumber();
            for (int i = lastColumnUsed; i > 0; i--)
            {
                var coluna = worksheet.Column(i);
                if (coluna.IsEmpty())
                {
                    coluna.Delete();
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Apply the alignment settings to the worksheet.
    /// <param name="worksheet">The worksheet to be processed.</param>
    /// <param name="verticalAlignment">The vertical alignment to be applied.</param>
    /// <param name="horizontalAlignment">The horizontal alignment to be applied.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool AlignmentCells(IXLWorksheet worksheet, string verticalAlignment, string horizontalAlignment)
    {
        try
        {
            switch (horizontalAlignment)
            {
                case "Left":
                    worksheet.Cells().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    break;
                case "Center":
                    worksheet.Cells().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    break;
                case "Right":
                    worksheet.Cells().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    break;
            }

            switch (verticalAlignment)
            {
                case "Top":
                    worksheet.Cells().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    break;
                case "Center":
                    worksheet.Cells().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    break;
                case "Bottom":
                    worksheet.Cells().Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    break;
            }

            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Apply the row height settings to the worksheet.
    /// <param name="worksheet">The worksheet to be processed.</param>
    /// <param name="rowHeight">The row height to be applied.</param>
    /// <param name="rowMaxHeight">The maximum row height to be applied.</param>
    /// <param name="auto">True to apply the auto row height, false to apply the row height.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool RowHeight(IXLWorksheet worksheet, double rowHeight, double rowMaxHeight, bool auto)
    {
        try
        {
            if (auto)
            {
                foreach (var row in worksheet.Rows())
                {
                    row.AdjustToContents();
                    if (row.Height > rowMaxHeight)
                    {
                        row.Height = rowMaxHeight;
                    }
                }
            }
            else
            {
                worksheet.Rows().Height = rowHeight;
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Apply the column width settings to the worksheet.
    /// <param name="worksheet">The worksheet to be processed.</param>
    /// <param name="columnWidth">The column width to be applied.</param>
    /// <param name="columnMaxWidth">The maximum column width to be applied.</param>
    /// <param name="auto">True to apply the auto column width, false to apply the column width.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool ColumnWidth(IXLWorksheet worksheet, double columnWidth, double columnMaxWidth, bool auto)
    {
        try
        {
            if (auto)
            {
                foreach (var column in worksheet.Columns())
                {
                    column.AdjustToContents();
                    if (column.Width > columnMaxWidth)
                    {
                        column.Width = columnMaxWidth;
                    }
                }
            }
            else
            {
                worksheet.Columns().Width = columnWidth;
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Apply the zoom settings to the worksheet.
    /// <param name="worksheet">The worksheet to be processed.</param>
    /// <param name="zoom">The zoom to be applied.</param>
    /// <returns>True if the operation was successful, false otherwise.</returns>
    /// </summary>
    public static bool Zoom(IXLWorksheet worksheet, int zoom)
    {
        try
        {
            worksheet.SheetView.ZoomScale = zoom;
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Copy the workbook to a new workbook.
    /// <param name="oldWorkbook">The workbook to be copied.</param>
    /// <param name="definitions">The definitions to be applied.</param>
    /// <returns>The new workbook.</returns>
    /// </summary>
    public static XLWorkbook? CopyWorkbook(XLWorkbook oldWorkbook, Definitions definitions)
    {
        try
        {
            var newWorkbook = new XLWorkbook();
            List<int> hiddenRows = new();

            foreach (var oldWorksheet in oldWorkbook.Worksheets)
            {
                hiddenRows.Clear();
                if (definitions.RemoveHiddenWorksheets)
                {
                    if (oldWorksheet.Visibility != XLWorksheetVisibility.Visible) { continue; } // Se a planilha não estiver visível, pula para a próxima
                }

                if (definitions.RemoveEmptyWorksheets)
                {
                    if (oldWorksheet.IsEmpty()) { continue; } // Se a planilha estiver vazia, pula para a próxima
                }

                string worksheetName = oldWorksheet.Name.ToUpperInvariant();
                newWorkbook.Worksheets.Add(worksheetName);

                var lastColumnUsed = oldWorksheet.LastColumnUsed().ColumnNumber();
                var lastRowUsed = oldWorksheet.LastRowUsed().RowNumber();

                if (definitions.GetLastRealEmptyRow)
                {
                    int rowEmptyCount = 0;
                    for (int row = 1; row <= lastRowUsed; row++)
                    {
                        if (oldWorksheet.Row(row).IsEmpty())
                        {
                            rowEmptyCount++;
                        }
                        else
                        {
                            rowEmptyCount = 0;
                        }

                        if (rowEmptyCount >= 10)
                        {
                            lastRowUsed = row - 10;
                            MessageBox.Show($"A linha [{lastRowUsed}] foi definida como a ultima linha da planilha [{worksheetName}!\nPor favor, verifique se essa informação está correta na planilha original.]", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                            break;
                        }
                    }
                }

                if (definitions.GetLastRealEmptyColumn)
                {
                    int columnEmptyCount = 0;
                    for (int column = 1; column <= lastColumnUsed; column++)
                    {
                        if (oldWorksheet.Column(column).IsEmpty())
                        {
                            columnEmptyCount++;
                        }
                        else
                        {
                            columnEmptyCount = 0;
                        }

                        if (columnEmptyCount >= 10)
                        {
                            lastColumnUsed = column - 10;
                            MessageBox.Show($"A coluna [{lastColumnUsed}] foi definida como a ultima coluna da planilha [{worksheetName}]!\nPor favor, verifique se essa informação está correta na planilha original.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                            break;
                        }
                    }
                }


                for (int row = 1; row <= lastRowUsed; row++)
                {
                    if (definitions.RemoveHiddenRows)
                    {
                        if (oldWorksheet.Row(row).IsHidden)
                        {
                            hiddenRows.Add(row);
                        }
                    }

                    for (int column = 1; column <= lastColumnUsed; column++)
                    {
                        if (!definitions.RemoveFormatting) // Se não for remover a formatação
                        {
                            newWorkbook.Worksheet(worksheetName).Cell(row, column).Style.NumberFormat = oldWorksheet.Cell(row, column).Style.NumberFormat;
                        }

                        if (!definitions.RemoveBackgroundColor)  // Se não for remover a cor de fundo
                        {
                            newWorkbook.Worksheet(worksheetName).Cell(row, column).Style.Fill.BackgroundColor = oldWorksheet.Cell(row, column).Style.Fill.BackgroundColor;
                        }

                        if (!definitions.RemoveFontColor) // Se não for remover a cor da fonte
                        {
                            newWorkbook.Worksheet(worksheetName).Cell(row, column).Style.Font.FontColor = oldWorksheet.Cell(row, column).Style.Font.FontColor;
                        }

                        if (definitions.FontSettings != null) // Se tiver fonte definida
                        {
                            newWorkbook.Worksheet(worksheetName).Cell(row, column).Style.Font.FontName = definitions.FontSettings.Name;
                            newWorkbook.Worksheet(worksheetName).Cell(row, column).Style.Font.FontSize = definitions.FontSettings.Size;
                        }

                        try
                        {
                            newWorkbook.Worksheet(worksheetName).Cell(row, column).Value = oldWorksheet.Cell(row, column).Value;
                        }
                        catch (Exception)
                        {
                            newWorkbook.Worksheet(worksheetName).Cell(row, column).Value = oldWorksheet.Cell(row, column).CachedValue;
                        }
                    }
                }

                hiddenRows = hiddenRows.OrderByDescending(row => row).ToList();
                foreach (var hiddenRow in hiddenRows)
                {
                    newWorkbook.Worksheet(worksheetName).Row(hiddenRow).Delete();
                }

            }

            return newWorkbook;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"{ex}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);


            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name}\nMessage: {ex.Message}",
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
            return null;
        }
    }

    /// <summary>
    /// Find the header of the worksheet.
    /// <param name="worksheet">The worksheet to be searched.</param>
    /// <param name="list">The list of strings to be searched.</param>
    /// <returns>True if the header was found, false if not.</returns>
    /// </summary>
    public static bool FindHeader(IXLWorksheet worksheet, List<string> list)
    {
        try
        {
            var lastRowUsed = worksheet.LastRowUsed().RowNumber();

            for (int row = 1; row <= lastRowUsed; row++)
            {
                var cellValues = worksheet.Row(row)
                                                          .Cells()
                                                          .Select(cell => Tools.ProcessTextForFindHeader(text: cell.CachedValue.ToString()));

                if (cellValues.Count(cellValue => list.Any(item => cellValue.Contains(item))) >= 3)
                {
                    int headerRow = row;

                    if (headerRow < 2) { return true; }

                    for (int i = headerRow - 1; i >= 1; i--)
                    {
                        worksheet.Row(i).Delete();
                    }
                    return true;
                }
            }
            return false;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Remove the columns depending on the keyword and filter.
    /// <param name="worksheet">The worksheet to be searched.</param>
    /// <param name="removeColumnsItems">The list of RemoveColumns objects.</param>
    /// <returns>True if the columns were removed, false if not.</returns>
    /// </summary>
    public static bool RemoveColumns(IXLWorksheet worksheet, List<RemoveColumns> removeColumnsItems)
    {
        try
        {
            if (removeColumnsItems is null) { return false; }

            var lastColumnUsed = worksheet.LastColumnUsed().ColumnNumber();

            for (int column = lastColumnUsed; column >= 1; column--)
            {
                var cellValue = worksheet.Column(column).Cell(1).CachedValue.ToString();
                cellValue = Tools.ProcessTextForRemoveColumns(text: cellValue);

                foreach (var removeColumnsItem in removeColumnsItems)
                {
                    if (removeColumnsItem.Keyword == null) { continue; }

                    var keyword = Tools.ProcessTextForRemoveColumns(text: removeColumnsItem.Keyword);

                    bool shouldDelete = false;

                    switch (removeColumnsItem.Filter)
                    {
                        case ERemoveColumnsFilterType.Equals:
                            shouldDelete = (cellValue == keyword);
                            break;
                        case ERemoveColumnsFilterType.Contains:
                            shouldDelete = cellValue.Contains(keyword);
                            break;
                        case ERemoveColumnsFilterType.StartsWith:
                            shouldDelete = cellValue.StartsWith(keyword);
                            break;
                        case ERemoveColumnsFilterType.EndsWith:
                            shouldDelete = cellValue.EndsWith(keyword);
                            break;
                    }

                    if (shouldDelete)
                    {
                        worksheet.Column(column).Delete();
                        break;
                    }
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Remove hidden rows.
    /// <param name="worksheet">The worksheet to be searched.</param>
    /// <returns>True if the rows were removed, false if not.</returns>
    /// </summary>
    public static bool RemoveHiddenRows(IXLWorksheet worksheet)
    {
        try
        {
            var lastRowUsed = worksheet.LastRowUsed().RowNumber();
            for (int i = lastRowUsed; i > 0; i--)
            {
                var linha = worksheet.Row(i);
                if (linha.IsHidden)
                {
                    linha.Delete();
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Rename the columns depending on the keyword and filter.
    /// <param name="worksheet">The worksheet to be searched.</param>
    /// <param name="renameColumnsItems">The list of RenameColumns objects.</param>
    /// <returns>True if the columns were renamed, false if not.</returns>
    /// </summary>
    public static bool RenameColumns(IXLWorksheet worksheet, List<RenameColumns> renameColumnsItems)
    {
        try
        {
            if (renameColumnsItems is null) { return false; }

            var lastColumnUsed = worksheet.LastColumnUsed().ColumnNumber();

            for (int column = lastColumnUsed; column >= 1; column--)
            {
                var cellValue = worksheet.Column(column).Cell(1).CachedValue.ToString();
                cellValue = Tools.ProcessTextForComparison(text: cellValue);

                foreach (var renameColumnsItem in renameColumnsItems)
                {
                    if (renameColumnsItem.Find == null) { continue; }

                    var find = Tools.ProcessTextForComparison(text: renameColumnsItem.Find);

                    bool shouldRename = false;

                    switch (renameColumnsItem.Filter)
                    {
                        case ERenameColumnsFilterType.Equals:
                            shouldRename = (cellValue == find);
                            break;
                        case ERenameColumnsFilterType.Contains:
                            shouldRename = cellValue.Contains(find);
                            break;
                        case ERenameColumnsFilterType.StartsWith:
                            shouldRename = cellValue.StartsWith(find);
                            break;
                        case ERenameColumnsFilterType.EndsWith:
                            shouldRename = cellValue.EndsWith(find);
                            break;
                    }

                    if (shouldRename)
                    {
                        worksheet.Column(column).Cell(1).Value = renameColumnsItem.Replace;
                        break;
                    }
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }
}
