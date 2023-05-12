using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;

namespace GCScriptExcelTools;

public static class ExcelFunctions
{
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

    public static bool RemoveFormulas(XLWorkbook workbook)
    {
        try
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var cell in worksheet.CellsUsed())
                {
                    cell.Value = cell.Value;
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
    public static bool RemoveBackgroundColor(IXLWorksheet worksheet)
    {
        try
        {
            worksheet.Cells().Style.Fill.BackgroundColor = XLColor.NoColor;
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    public static bool RemoveFontColor(IXLWorksheet worksheet)
    {
        try
        {
            worksheet.Cells().Style.Font.FontColor = XLColor.Black;
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    public static bool RemoveBorders(IXLWorksheet worksheet)
    {
        try
        {
            worksheet.Cells().Style.Border.OutsideBorder = XLBorderStyleValues.None;
            worksheet.Cells().Style.Border.InsideBorder = XLBorderStyleValues.None;
            worksheet.Cells().Style.Border.LeftBorder = XLBorderStyleValues.None;
            worksheet.Cells().Style.Border.RightBorder = XLBorderStyleValues.None;
            worksheet.Cells().Style.Border.TopBorder = XLBorderStyleValues.None;
            worksheet.Cells().Style.Border.BottomBorder = XLBorderStyleValues.None;
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    public static bool RemoveHyperlinks(IXLWorksheet worksheet)
    {
        try
        {
            foreach (var item in worksheet.Hyperlinks)
            {
                item.Delete();
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    public static bool RemoveConditionalFormatting(IXLWorksheet worksheet)
    {
        try
        {
            worksheet.ConditionalFormats.RemoveAll();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }
    public static bool SetFont(IXLWorksheet worksheet, string fontName, int fontSize)
    {
        try
        {
            worksheet.Cells().Style.Font.FontName = fontName;
            worksheet.Cells().Style.Font.FontSize = fontSize;
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }
    public static bool RemoveComments(IXLWorksheet worksheet)
    {
        try
        {
            worksheet.DeleteComments();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    public static bool RemoveMergeCells(IXLWorksheet worksheet)
    {
        try
        {
            foreach (var mergedRange in worksheet.MergedRanges)
            {
                worksheet.Range(mergedRange.RangeAddress).Unmerge();
            }
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

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

    public static bool RemoveFilter(IXLWorksheet worksheet)
    {
        try
        {
            worksheet.AutoFilter.Clear();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    public static bool RemoveWrapText(IXLWorksheet worksheet)
    {
        try
        {
            worksheet.Cells().Style.Alignment.WrapText = false;
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Method: {MethodBase.GetCurrentMethod()!.Name} | Message: {ex.Message}");
            return false;
        }
    }

    public static XLWorkbook? CopyWorkbook(XLWorkbook oldWorkbook, Definitions definitions)
    {
        try
        {
            var workbook = new XLWorkbook();
            foreach (var oldWorksheet in oldWorkbook.Worksheets)
            {
                if (definitions.RemoveInvisibleWorksheets)
                {
                    if (oldWorksheet.Visibility != XLWorksheetVisibility.Visible) { continue; } // Se a planilha não estiver visível, pula para a próxima
                }

                if (definitions.RemoveEmptyWorksheets)
                {
                    if (oldWorksheet.IsEmpty()) { continue; } // Se a planilha estiver vazia, pula para a próxima
                }

                string worksheetName = oldWorksheet.Name.ToUpperInvariant();
                workbook.Worksheets.Add(worksheetName);

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
                    for (int column = 1; column <= lastColumnUsed; column++)
                    {
                        if (!definitions.RemoveFormatting) // Se não for remover a formatação
                        {
                            workbook.Worksheet(worksheetName).Cell(row, column).Style.NumberFormat = oldWorksheet.Cell(row, column).Style.NumberFormat;
                        }

                        if (!definitions.RemoveBackgroundColor)  // Se não for remover a cor de fundo
                        {
                            workbook.Worksheet(worksheetName).Cell(row, column).Style.Fill.BackgroundColor = oldWorksheet.Cell(row, column).Style.Fill.BackgroundColor;
                        }

                        if (!definitions.RemoveFontColor) // Se não for remover a cor da fonte
                        {
                            workbook.Worksheet(worksheetName).Cell(row, column).Style.Font.FontColor = oldWorksheet.Cell(row, column).Style.Font.FontColor;
                        }

                        if (definitions.FontSettings != null) // Se tiver fonte definida
                        {
                            workbook.Worksheet(worksheetName).Cell(row, column).Style.Font.FontName = definitions.FontSettings.Name;
                            workbook.Worksheet(worksheetName).Cell(row, column).Style.Font.FontSize = definitions.FontSettings.Size;
                        }

                        try
                        {
                            workbook.Worksheet(worksheetName).Cell(row, column).Value = oldWorksheet.Cell(row, column).Value;
                        }
                        catch (Exception)
                        {
                            workbook.Worksheet(worksheetName).Cell(row, column).Value = oldWorksheet.Cell(row, column).CachedValue;
                        }
                    }
                }
            }

            return workbook;
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

    public static bool FindHeader(IXLWorksheet worksheet, List<string> list)
    {
        try
        {
            var lastRowUsed = worksheet.LastRowUsed().RowNumber();

            for (int row = 1; row <= lastRowUsed; row++)
            {
                var cellValues = worksheet.Row(row).Cells().Select(cell => Tools.TreatText(cell.CachedValue.ToString()));

                //var teste = cellValues.Count(cellValue => list.Any(item => cellValue.Contains(item)));

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

    public static bool RemoveSpecificColumns(IXLWorksheet worksheet, List<string> list)
    {
        try
        {
            var lastColumnUsed = worksheet.LastColumnUsed().ColumnNumber();

            for (int column = lastColumnUsed; column >= 1; column--)
            {
                var cellValue = worksheet.Column(column).Cell(1).CachedValue.ToString();
                cellValue = Tools.TreatText(cellValue);

                if (list.Any(item => cellValue.Contains(item)))
                {
                    worksheet.Column(column).Delete();
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
