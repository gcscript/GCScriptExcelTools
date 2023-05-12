using ClosedXML.Excel;
using GCScriptExcelTools.Models;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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

            if (fileDialog.FileNames.Length == 0) { return; } // Se nenhum arquivo for selecionado, sai do m�todo

            foreach (var fileName in fileDialog.FileNames)
            {
                if (!File.Exists(fileName)) { continue; } // Se o arquivo n�o existir, pula para o pr�ximo
                if (lst_FilesPath.Items.Contains(fileName)) { continue; } // Se o arquivo j� estiver na lista, pula para o pr�ximo

                var extension = Path.GetExtension(fileName);
                if (extension.ToLower() != ".xlsx" && extension.ToLower() != ".xlsm" && extension.ToLower() != ".xltx" && extension.ToLower() != ".xltm") { continue; } // Se a extens�o n�o for suportada, pula para o pr�ximo
                lst_FilesPath.Items.Add(fileName); // Adiciona o arquivo na lista
            }

            if (lst_FilesPath.Items.Count > 0) { lst_FilesPath.SelectedIndex = 0; }
        }

        private void btn_RemoveFile_Click(object sender, EventArgs e)
        {
            RemoveListboxItem(lst_FilesPath);
        }

        private void btn_Start_Click(object sender, EventArgs e)
        {
            string dateTime = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            foreach (var item in lst_FilesPath.Items)
            {
                if (!File.Exists(item.ToString())) { continue; } // Se o arquivo n�o existir, pula para o pr�ximo
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

                #region Find Header
                if (chk_FindHeader.Checked)
                {
                    definitions.FindHeader = new FindHeader()
                    {
                        Items = lst_FindHeader.Items.Cast<string>().ToList(),
                    };
                }
                #endregion

                #region Remove Specific Columns
                if (chk_RemoveSpecificColumns.Checked)
                {
                    definitions.RemoveSpecificColumns = new RemoveSpecificColumns()
                    {
                        Items = lst_RemoveSpecificColumns.Items.Cast<string>().ToList(),
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

                var workbook = ExcelFunctions.CopyWorkbook(workbookOld, definitions); // Transfere todo o conte�do da planilha antiga para a nova mantendo a formata��o

                if (workbook is null) { return; }

                if (chk_SortWorksheets.Checked) { ExcelFunctions.SortWorksheets(workbook, cmb_SortWorksheets.Text == "Ascending"); } // Ordena as planilhas

                foreach (var worksheet in workbook.Worksheets)
                {
                    if (chk_CellAlignments.Checked) { ExcelFunctions.AlignmentCells(worksheet, cmb_VerticalCellAlignments.Text, cmb_HorizontalCellAlignments.Text); } // Define o alinhamento
                    if (chk_Zoom.Checked) { ExcelFunctions.Zoom(worksheet, (int)nud_Zoom.Value); } // Define o zoom

                    #region Ajustes Finais
                    if (chk_RemoveEmptyRows.Checked) { ExcelFunctions.RemoveEmptyRows(worksheet); } // Remove as linhas vazias
                    if (chk_RemoveEmptyColumns.Checked) { ExcelFunctions.RemoveEmptyColumns(worksheet); } // Remove as colunas vazias
                    if (chk_FindHeader.Checked) { ExcelFunctions.FindHeader(worksheet, definitions.FindHeader.Items); }
                    if (chk_RemoveSpecificColumns.Checked) { ExcelFunctions.RemoveSpecificColumns(worksheet, definitions.RemoveSpecificColumns.Items); }
                    if (chk_RowHeight.Checked) { ExcelFunctions.RowHeight(worksheet, (double)nud_RowHeight.Value, (double)nud_RowMaxHeight.Value, chk_RowHeightAuto.Checked); } // Define a altura da linha
                    if (chk_ColumnWidth.Checked) { ExcelFunctions.ColumnWidth(worksheet, (double)nud_ColumnWidth.Value, (double)nud_ColumnMaxWidth.Value, chk_ColumnWidthAuto.Checked); } // Define a largura da coluna
                    #endregion

                    worksheet.SheetView.TopLeftCellAddress = worksheet.FirstCell().Address; // Scrolla para o topo da planilha
                    worksheet.SelectedRanges.RemoveAll(); // Remove todas as sele��es
                    worksheet.FirstCell().SetActive(); // Define a primeira c�lula como ativa
                    worksheet.TabSelected = false; // Define a planilha como n�o ativa
                }
                workbook.Worksheet(1).SetTabActive(); // Define a primeira planilha como ativa
                workbook.SaveAs(Path.Combine(filePath!, $"_{fileName}_{dateTime}.xlsx"));
                workbook.Dispose();
            }

            MessageBox.Show($"Processo finalizado com sucesso!", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
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

        private void chk_FindHeader_CheckedChanged(object sender, EventArgs e)
        {
            tlp_FindHeader.Enabled = chk_FindHeader.Checked;
            tlp_FindHeader.Visible = chk_FindHeader.Checked;
        }

        private void btn_FindHeaderAdd_Click(object sender, EventArgs e)
        {
            string item = Tools.TreatText(txt_FindHeader.Text);

            if (string.IsNullOrEmpty(item)) { return; }
            if (lst_FindHeader.Items.Contains(item)) { return; } // Se o arquivo j� estiver na lista,
            lst_FindHeader.Items.Add(item); // Adiciona o arquivo na lista

            if (lst_FindHeader.Items.Count > 0) { lst_FindHeader.SelectedIndex = 0; }
        }

        private void btn_FindHeaderRemove_Click(object sender, EventArgs e)
        {
            RemoveListboxItem(lst_FindHeader);
        }

        private void RemoveListboxItem(ListBox listBox)
        {
            if (listBox.SelectedItem != null) // Se tiver um item selecionado
            {
                int index = listBox.SelectedIndex; // Pega o index do item selecionado
                listBox.Items.RemoveAt(index); // Remove o item selecionado
                if (listBox.Items.Count > 0) // Se a lista n�o estiver vazia
                {
                    // Seleciona o pr�ximo item da lista
                    if (index > listBox.Items.Count - 1) { listBox.SelectedIndex = listBox.Items.Count - 1; }
                    else { listBox.SelectedIndex = index; }
                }
            }
        }

        private void chk_RemoveSpecificColumns_CheckedChanged(object sender, EventArgs e)
        {
            tlp_RemoveSpecificColumns.Enabled = chk_RemoveSpecificColumns.Checked;
            tlp_RemoveSpecificColumns.Visible = chk_RemoveSpecificColumns.Checked;
        }

        private void btn_RemoveSpecificColumnsAdd_Click(object sender, EventArgs e)
        {
            string item = Tools.TreatText(txt_RemoveSpecificColumns.Text);

            if (string.IsNullOrEmpty(item)) { return; }
            if (lst_RemoveSpecificColumns.Items.Contains(item)) { return; } // Se o arquivo j� estiver na lista,
            lst_RemoveSpecificColumns.Items.Add(item); // Adiciona o arquivo na lista

            if (lst_RemoveSpecificColumns.Items.Count > 0) { lst_RemoveSpecificColumns.SelectedIndex = 0; }
        }

        private void btn_RemoveSpecificColumnsRemove_Click(object sender, EventArgs e)
        {
            RemoveListboxItem(lst_RemoveSpecificColumns);
        }
    }

}