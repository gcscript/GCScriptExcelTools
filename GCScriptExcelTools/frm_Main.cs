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

        private void frm_Main_Load(object sender, EventArgs e)
        {
            cmb_FontName.SelectedIndex = 0;
            cmb_VerticalCellAlignments.SelectedIndex = 1;
            cmb_HorizontalCellAlignments.SelectedIndex = 0;
            cmb_SortWorksheets.SelectedIndex = 0;
            cmb_FindHeaderFilterOption.SelectedIndex = 0;
            cmb_RemoveColumnsFilterOption.SelectedIndex = 0;
            cmb_Preset.SelectedIndex = 0;

            SetActiveFunction(btn_Apply, pnl_Apply);

            StartDgvRemoveColumns();
        }

        public void StartDgvRemoveColumns()
        {
            SetDarkModeDataGridView(dgv_RemoveColumns);
            CreateColumnDataGridView(dgv_RemoveColumns, "Filter", "Filter", 0, DataGridViewAutoSizeColumnMode.None);
            CreateColumnDataGridView(dgv_RemoveColumns, "Keyword", "Keyword", 0, DataGridViewAutoSizeColumnMode.Fill);
            dgv_RemoveColumns.Rows.Add("Contains", "Admissão");
            dgv_RemoveColumns.Rows.Add("Contains", "Função");
            dgv_RemoveColumns.Rows.Add("Contains", "Salário");
            dgv_RemoveColumns.Rows.Add("Contains", "Sindicato");
            dgv_RemoveColumns.Rows.Add("Contains", "Situação");
            dgv_RemoveColumns.Rows.Add("Contains", "VR");
            dgv_RemoveColumns.Rows.Add("Equals", "VA");
            dgv_RemoveColumns.Rows.Add("StartsWith", "VA ");
            dgv_RemoveColumns.Rows.Add("EndsWith", " VA");
        }

        private static void SetDarkModeDataGridView(DataGridView dgv)
        {
            Color BackColor = Color.FromArgb(20, 20, 20);
            Color ForeColor = Color.LightGray;
            Color SelectionBackColor = Color.FromArgb(40, 40, 40);
            Color SelectionForeColor = Color.LightGray;
            Color GridColor = Color.FromArgb(40, 40, 40);
            Color ColumnHeadersBackColor = Color.DarkGray;
            Color ColumnHeadersForeColor = Color.Black;
            Color ColumnHeadersSelectionBackColor = Color.DarkGray;
            Color ColumnHeadersSelectionForeColor = Color.Black;

            dgv.DefaultCellStyle.Font = new Font("Consolas", 8);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Consolas", 9, FontStyle.Bold); // Define o estilo da fonte do cabeçalho
            dgv.RowHeadersVisible = false; // Define a coluna do cabeçalho como invisível
            dgv.BackgroundColor = BackColor; // Define a cor de fundo do DataGridView
            dgv.DefaultCellStyle.BackColor = BackColor; // Define a cor de fundo das células para uma cor escura
            dgv.DefaultCellStyle.ForeColor = ForeColor; // Define a cor do texto das células para branco
            dgv.DefaultCellStyle.SelectionBackColor = SelectionBackColor; // Define a cor de fundo das células selecionadas para uma cor escura
            dgv.DefaultCellStyle.SelectionForeColor = SelectionForeColor; // Define a cor do texto das células selecionadas para branco
            dgv.ColumnHeadersDefaultCellStyle.BackColor = ColumnHeadersBackColor; // Define a cor de fundo do cabeçalho para uma cor escura
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = ColumnHeadersForeColor; // Define a cor do texto do cabeçalho para branco
            dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = ColumnHeadersSelectionBackColor; // Define a cor de fundo da célula do cabeçalho quando selecionada para uma cor escura
            dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = ColumnHeadersSelectionForeColor; // Define a cor do texto da célula do cabeçalho quando selecionada para branco
            dgv.GridColor = GridColor; // Define a cor da borda do DataGridView
            dgv.EnableHeadersVisualStyles = false; // Define se o estilo padrão do cabeçalho será usado
            dgv.BorderStyle = BorderStyle.None; // Define se a borda do DataGridView será exibida
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.None; // Define o estilo da borda das células
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; // Define o estilo da borda do cabeçalho
            dgv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; // Define o estilo da borda do cabeçalho da linha
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect; // Define o modo de seleção
            dgv.MultiSelect = false; // Define se o usuário pode selecionar mais de uma linha
            dgv.AllowUserToAddRows = false; // Define se o usuário pode adicionar linhas
            dgv.AllowUserToDeleteRows = false; // Define se o usuário pode deletar linhas
            dgv.AllowUserToResizeRows = false; // Define se o usuário pode redimensionar as linhas
            dgv.AllowUserToOrderColumns = true; // Define se o usuário pode ordenar as colunas;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing; // Define se o usuário pode redimensionar a altura do cabeçalho
        }

        private void CreateColumnDataGridView(DataGridView dgv, string name, string header, int type, DataGridViewAutoSizeColumnMode autoSizeMode)
        {
            DataGridViewColumn column = new()
            {
                Name = name,
                HeaderText = header,
                ReadOnly = false,
                Visible = true,
                SortMode = DataGridViewColumnSortMode.Automatic,

                CellTemplate = type switch
                {
                    0 => new DataGridViewTextBoxCell(),
                    1 => new DataGridViewComboBoxCell(),
                    2 => new DataGridViewCheckBoxCell(),
                    _ => new DataGridViewTextBoxCell(),
                },

                AutoSizeMode = autoSizeMode,
            };

            dgv.Columns.Add(column);
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

                #region Remove Columns
                if (chk_RemoveColumns.Checked)
                {
                    definitions.RemoveColumnsItems = new();

                    foreach (DataGridViewRow removeColumnsRow in dgv_RemoveColumns.Rows)
                    {
                        if (removeColumnsRow.Cells["Filter"].Value == null || removeColumnsRow.Cells["Keyword"].Value == null) { continue; }

                        RemoveColumns removeColumnsItem = new()
                        {

                            Filter = Enum.TryParse(removeColumnsRow.Cells["Filter"].Value.ToString(),
                                                   out ERemoveColumnsFilterType filterType) ? filterType : throw new ArgumentException($"Invalid filter value: {removeColumnsRow.Cells["Filter"].Value}"),

                            Keyword = removeColumnsRow.Cells["Keyword"].Value.ToString()!
                        };

                        definitions.RemoveColumnsItems.Add(removeColumnsItem);
                    }

                }
                #endregion

                definitions.RemoveHiddenWorksheets = chk_RemoveInvisibleWorksheets.Checked;
                definitions.RemoveHiddenRows = chk_RemoveHiddenRows.Checked;
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
                    if (chk_RemoveColumns.Checked) { ExcelFunctions.RemoveColumns(worksheet, definitions.RemoveColumnsItems); }
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
            string item = Tools.ProcessTextForFindHeader(text: txt_FindHeader.Text);

            if (string.IsNullOrEmpty(item)) { return; }
            if (lst_FindHeader.Items.Contains(item)) { return; } // Se o arquivo ja estiver na lista,
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

        private void chk_RemoveColumns_CheckedChanged(object sender, EventArgs e)
        {
            tlp_RemoveColumns.Enabled = chk_RemoveColumns.Checked;
            tlp_RemoveColumns.Visible = chk_RemoveColumns.Checked;
        }

        private void btn_RemoveColumnsAdd_Click(object sender, EventArgs e)
        {
            string originalKeyword = txt_RemoveColumns.Text;
            string originalFilter = cmb_RemoveColumnsFilterOption.Text;
            string processedKeyword = Tools.ProcessTextForRemoveColumns(originalKeyword);

            if (string.IsNullOrEmpty(processedKeyword)) { return; }

            foreach (DataGridViewRow row in dgv_RemoveColumns.Rows)
            {
                string columnNameValue = Tools.ProcessTextForRemoveColumns(text: row.Cells["Keyword"].Value.ToString()!);

                if (columnNameValue == processedKeyword)
                {
                    txt_RemoveColumns.Text = string.Empty;
                    MessageBox.Show($"[{originalKeyword}] already exists!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            var index = dgv_RemoveColumns.Rows.Add();
            dgv_RemoveColumns.Rows[index].Cells["Filter"].Value = originalFilter;
            dgv_RemoveColumns.Rows[index].Cells["Keyword"].Value = originalKeyword;

            txt_RemoveColumns.Text = string.Empty;
            txt_RemoveColumns.Focus();
        }

        private void btn_RemoveColumnsRemove_Click(object sender, EventArgs e)
        {
            if (dgv_RemoveColumns.SelectedRows.Count > 0)
            {
                dgv_RemoveColumns.Rows.RemoveAt(dgv_RemoveColumns.SelectedRows[0].Index);
            }
        }

        private void cmb_Preset_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}