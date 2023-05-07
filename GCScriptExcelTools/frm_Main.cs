using ClosedXML.Excel;
using GCScriptExcelTools.Models;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace GCScriptExcelTools
{
    public partial class frm_Main : Form
    {
        public frm_Main()
        {
            InitializeComponent();
        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            FileDialog fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xltx;*.xltm"
            };
            fileDialog.ShowDialog();

            if (fileDialog.FileNames.Length == 0) { return; } // Se nenhum arquivo for selecionado, sai do método

            foreach (var fileName in fileDialog.FileNames)
            {
                if (!File.Exists(fileName)) { continue; } // Se o arquivo não existir, pula para o próximo
                if (lst_FilesPath.Items.Contains(fileName)) { continue; } // Se o arquivo já estiver na lista, pula para o próximo

                var extension = Path.GetExtension(fileName);
                if (extension != ".xlsx" && extension != ".xlsm" && extension != ".xltx" && extension != ".xltm") { continue; } // Se a extensão não for suportada, pula para o próximo
                lst_FilesPath.Items.Add(fileName); // Adiciona o arquivo na lista
            }

            if (lst_FilesPath.Items.Count > 0) { lst_FilesPath.SelectedIndex = 0; }
        }

        private void btn_RemoveFile_Click(object sender, EventArgs e)
        {
            if (lst_FilesPath.SelectedItem != null) // Se tiver um item selecionado
            {
                int index = lst_FilesPath.SelectedIndex; // Pega o index do item selecionado
                lst_FilesPath.Items.RemoveAt(index); // Remove o item selecionado
                if (lst_FilesPath.Items.Count > 0) // Se a lista não estiver vazia
                {
                    // Seleciona o próximo item da lista
                    if (index > lst_FilesPath.Items.Count - 1) { lst_FilesPath.SelectedIndex = lst_FilesPath.Items.Count - 1; }
                    else { lst_FilesPath.SelectedIndex = index; }
                }
            }
        }

        private void btn_Start_Click(object sender, EventArgs e)
        {
            foreach (var item in lst_FilesPath.Items)
            {
                if (!File.Exists(item.ToString())) { continue; } // Se o arquivo não existir, pula para o próximo
                var workbookOld = new XLWorkbook(item.ToString());
                var fileName = Path.GetFileNameWithoutExtension(item.ToString());
                var filePath = Path.GetDirectoryName(item.ToString());

                Definitions definitions = new();

                #region Sort Worksheets
                if (chk_SortWorksheets.Checked)
                {
                    definitions.SortWorksheets = new SortWorksheets()
                    {
                        By = cmb_SortWorksheets.Text
                    };
                }
                #endregion

                #region Font Settings
                if (chk_Font.Checked)
                {
                    definitions.FontSettings = new FontSettings()
                    {
                        Name = cmb_FontName.Text,
                        Size = (int)nud_FontSize.Value
                    };
                }
                #endregion

                definitions.RemoveInvisibleWorksheets = chk_RemoveInvisibleWorksheets.Checked;
                definitions.RemoveEmptyWorksheets = chk_RemoveEmptyWorksheets.Checked;
                definitions.RemoveFormatting = chk_RemoveFormatting.Checked;
                definitions.RemoveBackgroundColor = chk_RemoveBackgroundColor.Checked;
                definitions.RemoveFontColor = chk_RemoveFontColor.Checked;
                definitions.GetLastRealEmptyRow = chk_GetLastRealEmptyRow.Checked;
                definitions.GetLastRealEmptyColumn = chk_GetLastRealEmptyColumn.Checked;

                var workbook = ExcelFunctions.CopyWorkbook(workbookOld, definitions); // Transfere todo o conteúdo da planilha antiga para a nova mantendo a formatação

                if (workbook is null) { return; }

                if (chk_SortWorksheets.Checked) { ExcelFunctions.SortWorksheets(workbook, cmb_SortWorksheets.Text == "Ascending"); } // Ordena as planilhas

                foreach (var worksheet in workbook.Worksheets)
                {
                    if (chk_CellAlignments.Checked) { ExcelFunctions.AlignmentCells(worksheet, cmb_VerticalCellAlignments.Text, cmb_HorizontalCellAlignments.Text); } // Define o alinhamento
                    if (chk_Zoom.Checked) { ExcelFunctions.Zoom(worksheet, (int)nud_Zoom.Value); } // Define o zoom

                    #region Ajustes Finais
                    if (chk_RowHeight.Checked) { ExcelFunctions.RowHeight(worksheet, (double)nud_RowHeight.Value, (double)nud_RowMaxHeight.Value, chk_RowHeightAuto.Checked); } // Define a altura da linha
                    if (chk_ColumnWidth.Checked) { ExcelFunctions.ColumnWidth(worksheet, (double)nud_ColumnWidth.Value, (double)nud_ColumnMaxWidth.Value, chk_ColumnWidthAuto.Checked); } // Define a largura da coluna
                    if (chk_RemoveEmptyRows.Checked) { ExcelFunctions.RemoveEmptyRows(worksheet); } // Remove as linhas vazias
                    if (chk_RemoveEmptyColumns.Checked) { ExcelFunctions.RemoveEmptyColumns(worksheet); } // Remove as colunas vazias
                    #endregion

                    worksheet.SheetView.TopLeftCellAddress = worksheet.FirstCell().Address; // Scrolla para o topo da planilha
                    worksheet.SelectedRanges.RemoveAll(); // Remove todas as seleções
                    worksheet.FirstCell().SetActive(); // Define a primeira célula como ativa
                    worksheet.TabSelected = false; // Define a planilha como não ativa
                }
                workbook.Worksheet(1).SetTabActive(); // Define a primeira planilha como ativa
                workbook.SaveAs(Path.Combine(filePath!, $"_{fileName}_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx"));
                workbook.Dispose();
            }

            MessageBox.Show("Fim");
        }

        private void btn_Start_Click_Backup(object sender, EventArgs e)
        {
            //foreach (var item in lst_FilesPath.Items)
            //{
            //    if (!File.Exists(item.ToString())) { continue; } // Se o arquivo não existir, pula para o próximo
            //    var workbook = new XLWorkbook(item.ToString());
            //    var fileName = Path.GetFileNameWithoutExtension(item.ToString());
            //    var filePath = Path.GetDirectoryName(item.ToString());

            //    if (chk_RemoveFormatting.Checked) { ExcelFunctions.RemoveFormulas(workbook); } // Remove as fórmulas
            //    if (chk_RemoveEmptyWorksheets.Checked) { ExcelFunctions.RemoveEmptyWorksheets(workbook); } // Remove as planilhas vazias
            //    if (chk_SortWorksheets.Checked) { ExcelFunctions.SortWorksheets(workbook, cmb_SortWorksheets.Text == "Ascending"); } // Ordena as planilhas

            //    foreach (var worksheet in workbook.Worksheets)
            //    {
            //        if (worksheet.Visibility != XLWorksheetVisibility.Visible) { continue; } // Se a planilha não estiver visível, pula para a próxima

            //        #region Ajustes Iniciais
            //        if (chk_RemoveFilter.Checked) { ExcelFunctions.RemoveFilter(worksheet); } // Remove os filtros
            //        if (chk_RemoveHyperlinks.Checked) { ExcelFunctions.RemoveHyperlinks(worksheet); } // Remove os hiperlinks
            //        if (chk_RemoveConditionalFormatting.Checked) { ExcelFunctions.RemoveConditionalFormatting(worksheet); } // Remove a formatação condicional
            //        if (chk_RemoveWrapText.Checked) { ExcelFunctions.RemoveWrapText(worksheet); } // Remove Quebras de Linha
            //        #endregion

            //        if (chk_RemoveFontColor.Checked) { ExcelFunctions.RemoveFontColor(worksheet); } // Remove a cor da fonte
            //        if (chk_RemoveBackgroundColor.Checked) { ExcelFunctions.RemoveBackgroundColor(worksheet); } // Remove a cor de fundo
            //        if (chk_RemoveBorders.Checked) { ExcelFunctions.RemoveBorders(worksheet); } // Remove as bordas
            //        if (chk_Font.Checked)
            //        {
            //            ExcelFunctions.SetFont(worksheet,
            //                                           cmb_FontName.Text,
            //                                           (int)nud_FontSize.Value);
            //        } // Define as configurações da fonte

            //        if (chk_RemoveComments.Checked) { ExcelFunctions.RemoveComments(worksheet); } // Remove os comentários
            //        if (chk_RemoveMergeCells.Checked) { ExcelFunctions.RemoveMergeCells(worksheet); } // Remove as células mescladas
            //        if (chk_CellAlignments.Checked) { ExcelFunctions.AlignmentCells(worksheet, cmb_VerticalCellAlignments.Text, cmb_HorizontalCellAlignments.Text); } // Define o alinhamento
            //        if (chk_Zoom.Checked) { ExcelFunctions.Zoom(worksheet, (int)nud_Zoom.Value); } // Define o zoom

            //        #region Ajustes Finais
            //        if (chk_RowHeight.Checked) { ExcelFunctions.RowHeight(worksheet, (double)nud_RowHeight.Value, (double)nud_RowMaxHeight.Value); } // Define a altura da linha
            //        if (chk_ColumnWidth.Checked) { ExcelFunctions.ColumnWidth(worksheet, (double)nud_ColumnWidth.Value, (double)nud_ColumnMaxWidth.Value); } // Define a largura da coluna
            //        if (chk_RemoveEmptyRows.Checked) { ExcelFunctions.RemoveEmptyRows(worksheet); } // Remove as linhas vazias
            //        if (chk_RemoveEmptyColumns.Checked) { ExcelFunctions.RemoveEmptyColumns(worksheet); } // Remove as colunas vazias
            //        if (chk_RemovePictures.Checked) { ExcelFunctions.RemovePictures(worksheet); } // Remove as imagens
            //        #endregion

            //        worksheet.SheetView.TopLeftCellAddress = worksheet.FirstCell().Address; // Scrolla para o topo da planilha
            //        worksheet.SelectedRanges.RemoveAll(); // Remove todas as seleções
            //        worksheet.FirstCell().SetActive(); // Define a primeira célula como ativa
            //        worksheet.TabSelected = false; // Define a planilha como não ativa
            //    }

            //    workbook.Worksheet(1).SetTabActive(); // Define a primeira planilha como ativa

            //    workbook.SaveAs(Path.Combine(filePath!, $"_{fileName}_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx"));
            //    workbook.Dispose();
            //}

            //MessageBox.Show("Fim");
        }

        private void frm_Main_Load(object sender, EventArgs e)
        {
            cmb_FontName.SelectedIndex = 0;
            cmb_VerticalCellAlignments.SelectedIndex = 1;
            cmb_HorizontalCellAlignments.SelectedIndex = 0;
            cmb_SortWorksheets.SelectedIndex = 0;

            SetActiveFunction(btn_Apply, pnl_Apply);

        }

        private void chk_FontSettings_CheckedChanged(object sender, EventArgs e)
        {
            tlp_Font.Enabled = chk_Font.Checked;
            tlp_Font.Visible = chk_Font.Checked;
        }

        private void chk_CellAlignments_CheckedChanged(object sender, EventArgs e)
        {
            tlp_CellAlignments.Enabled = chk_CellAlignments.Checked;
            tlp_CellAlignments.Visible = chk_CellAlignments.Checked;
        }

        private void chk_RowHeight_CheckedChanged(object sender, EventArgs e)
        {
            tlp_RowHeight.Enabled = chk_RowHeight.Checked;
            tlp_RowHeight.Visible = chk_RowHeight.Checked;
        }

        private void chk_ColumnWidth_CheckedChanged(object sender, EventArgs e)
        {
            tlp_ColumnWidth.Enabled = chk_ColumnWidth.Checked;
            tlp_ColumnWidth.Visible = chk_ColumnWidth.Checked;
        }

        private void chk_SortWorksheets_CheckedChanged(object sender, EventArgs e)
        {
            tlp_SortWorksheets.Enabled = chk_SortWorksheets.Checked;
            tlp_SortWorksheets.Visible = chk_SortWorksheets.Checked;
        }

        private void chk_Zoom_CheckedChanged(object sender, EventArgs e)
        {
            tlp_Zoom.Enabled = chk_Zoom.Checked;
            tlp_Zoom.Visible = chk_Zoom.Checked;

        }

        private void btn_Apply_Click(object sender, EventArgs e)
        {
            SetActiveFunction(btn_Apply, pnl_Apply);
        }

        private void btn_Remove_Click(object sender, EventArgs e)
        {
            SetActiveFunction(btn_Remove, pnl_Remove);
        }

        private void btn_Others_Click(object sender, EventArgs e)
        {
            SetActiveFunction(btn_Others, pnl_Others);
        }

        private void HidePanels(Panel[] panels)
        {
            foreach (var panel in panels) { panel.Visible = false; }
            ResetButtons(new Button[] { btn_Apply, btn_Remove, btn_Others });
        }
        private void ResetButtons(Button[] buttons)
        {
            foreach (var button in buttons) { button.BackColor = System.Drawing.Color.FromArgb(20, 20, 20); }
        }

        private void SetActiveFunction(Button button, Panel panel)
        {
            HidePanels(new Panel[] { pnl_Apply, pnl_Remove, pnl_Others });
            panel.Visible = true;
            button.BackColor = System.Drawing.Color.FromArgb(80, 80, 80);
        }

        private void chk_RowHeightAuto_CheckedChanged(object sender, EventArgs e)
        {
            nud_RowHeight.Enabled = !chk_RowHeightAuto.Checked;
            nud_RowMaxHeight.Enabled = chk_RowHeightAuto.Checked;
        }

        private void chk_ColumnWidthAuto_CheckedChanged(object sender, EventArgs e)
        {
            nud_ColumnWidth.Enabled = !chk_ColumnWidthAuto.Checked;
            nud_ColumnMaxWidth.Enabled = chk_ColumnWidthAuto.Checked;
        }

        private void lst_FilesPath_SelectedIndexChanged(object sender, EventArgs e) => btn_Start.Enabled = lst_FilesPath.Items.Count > 0;

        private void lbl_Credits_Click(object sender, EventArgs e)
        {
            OpenUrl("https://gcscript.dev");
        }

        private void OpenUrl(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch
            {
                // hack because of this: https://github.com/dotnet/corefx/issues/10361
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    throw;
                }
            }
        }
    }
}