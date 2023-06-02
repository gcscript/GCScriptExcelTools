using ClosedXML.Excel;
using GCScriptExcelTools.Models;
using System.Diagnostics;
using System.Runtime.InteropServices;
using GCScriptExcelTools.Services;

namespace GCScriptExcelTools;

public partial class frm_MainOld : Form
{
    public frm_MainOld()
    {
        InitializeComponent();
        RemoveColumnsComponent();
        RenameColumnsComponent();
        SortColumnsComponent();
    }

    private void frm_Main_Load(object sender, EventArgs e)
    {
        cmb_FontName.SelectedIndex = 0;
        cmb_VerticalCellAlignments.SelectedIndex = 1;
        cmb_HorizontalCellAlignments.SelectedIndex = 0;
        cmb_SortWorksheets.SelectedIndex = 0;
        cmb_FindHeaderFilterOption.SelectedIndex = 0;
        cmb_RemoveColumnsFilterOption.SelectedIndex = 0;
        cmb_RenameColumnsFilterOption.SelectedIndex = 0;
        cmb_Preset.SelectedIndex = 0;

        SetActiveFunction(btn_Apply, pnl_Apply);

        LoadPresetList();
        //CreateSortColumnsComponent();
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
                definitions.RemoveColumnsList = new();

                foreach (DataGridViewRow removeColumnsRow in dgv_RemoveColumns.Rows)
                {
                    if (removeColumnsRow.Cells["Filter"].Value == null || removeColumnsRow.Cells["Keyword"].Value == null) { continue; }

                    RemoveColumns removeColumnsItem = new()
                    {

                        Filter = Enum.TryParse(removeColumnsRow.Cells["Filter"].Value.ToString(),
                                               out ERemoveColumnsFilterType filterType) ? filterType : throw new ArgumentException($"Invalid filter value: {removeColumnsRow.Cells["Filter"].Value}"),

                        Keyword = removeColumnsRow.Cells["Keyword"].Value.ToString()!
                    };

                    definitions.RemoveColumnsList.Add(removeColumnsItem);
                }

            }
            #endregion

            definitions.RemoveHiddenWorksheets = chk_RemoveHiddenWorksheets.Checked;
            definitions.RemoveHiddenRows = chk_RemoveHiddenRows.Checked;
            definitions.RemoveEmptyWorksheets = chk_RemoveEmptyWorksheets.Checked;
            definitions.RemoveFormatting = chk_RemoveFormatting.Checked;
            definitions.RemoveBackgroundColor = chk_RemoveBackgroundColor.Checked;
            definitions.RemoveFontColor = chk_RemoveFontColor.Checked;
            definitions.GetLastRealEmptyRow = chk_GetLastRealEmptyRow.Checked;
            definitions.GetLastRealEmptyColumn = chk_GetLastRealEmptyColumn.Checked;

            #region Rename Columns
            if (chk_RenameColumns.Checked)
            {
                definitions.RenameColumnsList = new();

                foreach (DataGridViewRow renameColumnsRow in dgv_RenameColumns.Rows)
                {
                    if (renameColumnsRow.Cells["Filter"].Value == null || renameColumnsRow.Cells["Find"].Value == null || renameColumnsRow.Cells["Replace"].Value == null) { continue; }

                    RenameColumns renameColumnsItem = new()
                    {

                        Filter = Enum.TryParse(renameColumnsRow.Cells["Filter"].Value.ToString(),
                                               out ERenameColumnsFilterType filterType) ? filterType : throw new ArgumentException($"Invalid filter value: {renameColumnsRow.Cells["Filter"].Value}"),

                        Find = renameColumnsRow.Cells["Find"].Value.ToString()!,
                        Replace = renameColumnsRow.Cells["Replace"].Value.ToString()!
                    };

                    definitions.RenameColumnsList.Add(renameColumnsItem);
                }

            }
            #endregion

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
                if (chk_RemoveColumns.Checked) { ExcelFunctions.RemoveColumns(worksheet, definitions.RemoveColumnsList); }
                if (chk_RenameColumns.Checked) { ExcelFunctions.RenameColumns(worksheet, definitions.RenameColumnsList); }
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

    private void cmb_Preset_SelectedIndexChanged(object sender, EventArgs e)
    {
        LoadPreset(cmb_Preset.Text);
    }

    private void btn_SavePreset_Click(object sender, EventArgs e)
    {
        tlp_SavePreset.Visible = true;
        tlp_SavePreset.Enabled = true;
        tlp_SelectPreset.Visible = false;
        tlp_SelectPreset.Enabled = false;
        txt_Preset.Select();
    }

    private void btn_ReplacePreset_Click(object sender, EventArgs e)
    {
        if (cmb_Preset.Text == "Default") { return; }
        SavePreset(cmb_Preset.Text);
    }

    private void btn_RemovePreset_Click(object sender, EventArgs e)
    {
        if (cmb_Preset.Text == "Default") { return; }
        RemovePreset(cmb_Preset.Text);
    }

    private void btn_ConfirmPreset_Click(object sender, EventArgs e)
    {
        if (SavePreset(txt_Preset.Text))
        {
            txt_Preset.Text = string.Empty;
            tlp_SelectPreset.Visible = true;
            tlp_SelectPreset.Enabled = true;
            tlp_SavePreset.Visible = false;
            tlp_SavePreset.Enabled = false;
            cmb_Preset.Select();
        }
    }

    private void btn_CancelPreset_Click(object sender, EventArgs e)
    {
        tlp_SelectPreset.Visible = true;
        tlp_SelectPreset.Enabled = true;
        tlp_SavePreset.Visible = false;
        tlp_SavePreset.Enabled = false;
        cmb_Preset.Select();
    }

    private bool SavePreset(string title)
    {
        try
        {
            if (string.IsNullOrEmpty(title)) { throw new Exception("Title is empty!"); }
            if (title == "Default") { return false; }
            Preset preset = new() { Title = title };
            preset.Remove.HiddenWorksheets = chk_RemoveHiddenWorksheets.Checked;
            preset.Remove.EmptyWorksheets = chk_RemoveEmptyWorksheets.Checked;
            preset.Remove.EmptyRows = chk_RemoveEmptyRows.Checked;
            preset.Remove.EmptyColumns = chk_RemoveEmptyColumns.Checked;
            preset.Remove.HiddenRows = chk_RemoveHiddenRows.Checked;
            preset.Remove.FontColor = chk_RemoveFontColor.Checked;
            preset.Remove.BackgroundColor = chk_RemoveBackgroundColor.Checked;
            preset.Remove.Formatting = chk_RemoveFormatting.Checked;

            List<RemoveColumns> removeColumns = new();

            foreach (DataGridViewRow removeColumnsRow in dgv_RemoveColumns.Rows)
            {
                if (removeColumnsRow.Cells["Filter"].Value == null || removeColumnsRow.Cells["Keyword"].Value == null) { continue; }

                RemoveColumns removeColumnsItem = new()
                {
                    Filter = Enum.TryParse(removeColumnsRow.Cells["Filter"].Value.ToString(),
                                           out ERemoveColumnsFilterType filterType) ? filterType : throw new ArgumentException($"Invalid filter value: {removeColumnsRow.Cells["Filter"].Value}"),

                    Keyword = removeColumnsRow.Cells["Keyword"].Value.ToString()!
                };

                removeColumns.Add(removeColumnsItem);
            }

            preset.Remove.Columns = removeColumns;

            preset.Others.GetLastRealEmptyRow = chk_GetLastRealEmptyRow.Checked;
            preset.Others.GetLastRealEmptyColumn = chk_GetLastRealEmptyColumn.Checked;

            preset.Others.FindHeader = new FindHeader() { Items = lst_FindHeader.Items.Cast<string>().ToList() };

            List<RenameColumns> renameColumns = new();

            foreach (DataGridViewRow renameColumnsRow in dgv_RenameColumns.Rows)
            {
                if (renameColumnsRow.Cells["Filter"].Value == null || renameColumnsRow.Cells["Find"].Value == null || renameColumnsRow.Cells["Replace"].Value == null) { continue; }

                RenameColumns renameColumnsItem = new()
                {
                    Filter = Enum.TryParse(renameColumnsRow.Cells["Filter"].Value.ToString(),
                                           out ERenameColumnsFilterType filterType) ? filterType : throw new ArgumentException($"Invalid filter value: {renameColumnsRow.Cells["Filter"].Value}"),

                    Find = renameColumnsRow.Cells["Find"].Value.ToString()!,
                    Replace = renameColumnsRow.Cells["Replace"].Value.ToString()!,
                };

                renameColumns.Add(renameColumnsItem);
            }

            preset.Others.RenameColumns = renameColumns;

            if (FPreset.Save(preset))
            {
                LoadPresetList(title);
                MessageBox.Show($"Preset [{title}] saved!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            else
            {
                throw new Exception($"Preset [{title}] not saved!");
                return false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    private void LoadPreset(string title)
    {
        try
        {
            if (string.IsNullOrEmpty(title)) { throw new Exception("Title is empty!"); }

            Preset preset = new();

            if (title == "Default")
            {
                preset.Remove.HiddenWorksheets = true;
                preset.Remove.EmptyWorksheets = true;
                preset.Remove.EmptyRows = true;
                preset.Remove.EmptyColumns = true;
                preset.Remove.HiddenRows = true;
                preset.Remove.FontColor = true;
                preset.Remove.BackgroundColor = true;
                preset.Remove.Formatting = false;

                preset.Others.GetLastRealEmptyRow = true;
                preset.Others.GetLastRealEmptyColumn = true;
            }
            else
            {
                preset = FPreset.Load(Path.Combine(Settings.PresetsPath, $"{title}.json")) ?? throw new Exception($"Preset [{title}] not found!");
            }

            chk_RemoveHiddenWorksheets.Checked = preset.Remove.HiddenWorksheets;
            chk_RemoveEmptyWorksheets.Checked = preset.Remove.EmptyWorksheets;
            chk_RemoveEmptyRows.Checked = preset.Remove.EmptyRows;
            chk_RemoveEmptyColumns.Checked = preset.Remove.EmptyColumns;
            chk_RemoveHiddenRows.Checked = preset.Remove.HiddenRows;
            chk_RemoveFontColor.Checked = preset.Remove.FontColor;
            chk_RemoveBackgroundColor.Checked = preset.Remove.BackgroundColor;
            chk_RemoveFormatting.Checked = preset.Remove.Formatting;

            dgv_RemoveColumns.Rows.Clear();

            if (preset.Remove.Columns is null)
            {
                chk_RemoveColumns.Checked = false;
            }
            else
            {
                chk_RemoveColumns.Checked = true;
                foreach (RemoveColumns removeColumns in preset.Remove.Columns)
                {
                    var index = dgv_RemoveColumns.Rows.Add();
                    dgv_RemoveColumns.Rows[index].Cells["Filter"].Value = removeColumns.Filter;
                    dgv_RemoveColumns.Rows[index].Cells["Keyword"].Value = removeColumns.Keyword;
                }
            }

            chk_GetLastRealEmptyRow.Checked = preset.Others.GetLastRealEmptyRow;
            chk_GetLastRealEmptyColumn.Checked = preset.Others.GetLastRealEmptyColumn;

            lst_FindHeader.Items.Clear();
            if (preset.Others.FindHeader is null)
            {
                chk_FindHeader.Checked = false;
            }
            else
            {
                chk_FindHeader.Checked = true;
                lst_FindHeader.Items.AddRange(preset.Others.FindHeader.Items.ToArray());
            }

            dgv_RenameColumns.Rows.Clear();

            if (preset.Others.RenameColumns is null)
            {
                chk_RenameColumns.Checked = false;
            }
            else
            {
                chk_RenameColumns.Checked = true;
                foreach (RenameColumns renameColumns in preset.Others.RenameColumns)
                {
                    var index = dgv_RenameColumns.Rows.Add();
                    dgv_RenameColumns.Rows[index].Cells["Filter"].Value = renameColumns.Filter;
                    dgv_RenameColumns.Rows[index].Cells["Find"].Value = renameColumns.Find;
                    dgv_RenameColumns.Rows[index].Cells["Replace"].Value = renameColumns.Replace;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private bool RemovePreset(string title)
    {
        try
        {
            cmb_Preset.Items.Remove(title);

            if (FPreset.Remove(Path.Combine(Settings.PresetsPath, $"{title}.json")))
            {
                LoadPresetList();
                MessageBox.Show($"Preset [{title}] removed!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }

            MessageBox.Show($"Preset [{title}] not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    private void LoadPresetList(string? selectedPreset = null)
    {
        try
        {
            cmb_Preset.Items.Clear();
            cmb_Preset.Items.Add("Default");
            List<string> presets = new();
            DirectoryInfo directoryInfo = new(Settings.PresetsPath);
            foreach (FileInfo fileInfo in directoryInfo.GetFiles("*.json"))
            {
                presets.Add(Path.GetFileNameWithoutExtension(fileInfo.Name));
            }
            presets.Sort();
            cmb_Preset.Items.AddRange(presets.ToArray());

            if (selectedPreset is not null)
            {
                cmb_Preset.SelectedItem = selectedPreset;
            }
            else
            {
                cmb_Preset.SelectedIndex = 0;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void btn_RenameColumnsAdd_Click(object sender, EventArgs e)
    {
        var filter = cmb_RenameColumnsFilterOption.Text;
        var find = txt_RenameColumnsFind.Text;
        var replace = txt_RenameColumnsReplace.Text;

        if (string.IsNullOrEmpty(find) || string.IsNullOrWhiteSpace(find))
        {
            MessageBox.Show(text: "Find cannot be empty!",
                            caption: "Error",
                            buttons: MessageBoxButtons.OK,
                            icon: MessageBoxIcon.Error);
            txt_RenameColumnsFind.Focus();
            return;
        }

        if (string.IsNullOrEmpty(replace) || string.IsNullOrWhiteSpace(replace))
        {
            MessageBox.Show(text: "Replace cannot be empty!",
                            caption: "Error",
                            buttons: MessageBoxButtons.OK,
                            icon: MessageBoxIcon.Error);
            txt_RenameColumnsReplace.Focus();
            return;
        }

        if (ComponentSettings.DataGridViewCheckIfTheValueExists(dgv_RenameColumns, "Find", find))
        {
            MessageBox.Show(text: $"\"{find}\" already exists!",
                            caption: "Error",
                            buttons: MessageBoxButtons.OK,
                            icon: MessageBoxIcon.Error);
            return;
        }

        var index = dgv_RenameColumns.Rows.Add();
        dgv_RenameColumns.Rows[index].Cells["Filter"].Value = filter;
        dgv_RenameColumns.Rows[index].Cells["Find"].Value = find;
        dgv_RenameColumns.Rows[index].Cells["Replace"].Value = replace;

        txt_RenameColumnsFind.Text = string.Empty;
        txt_RenameColumnsReplace.Text = string.Empty;
        txt_RenameColumnsFind.Focus();
    }

    private void btn_RenameColumnsRemove_Click(object sender, EventArgs e)
    {
        if (dgv_RenameColumns.SelectedRows.Count > 0)
        {
            dgv_RenameColumns.Rows.RemoveAt(dgv_RenameColumns.SelectedRows[0].Index);
        }
    }

    private void chk_RenameColumns_CheckedChanged(object sender, EventArgs e)
    {
        tlp_RenameColumns.Enabled = chk_RenameColumns.Checked;
        tlp_RenameColumns.Visible = chk_RenameColumns.Checked;
    }
}