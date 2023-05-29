namespace GCScriptExcelTools;
public partial class frm_Main
{
    private void Chk_SortColumns_CheckedChanged(object? sender, EventArgs e)
    {
        tlp_SortColumns.Enabled = chk_SortColumns.Checked;
        tlp_SortColumns.Visible = chk_SortColumns.Checked;
    }

    private void Btn_SortColumnsAdd_Click(object? sender, EventArgs e)
    {
        var keyword = txt_SortColumns.Text;

        if (string.IsNullOrEmpty(keyword) || string.IsNullOrWhiteSpace(keyword))
        {
            MessageBox.Show("Keyword cannot be empty!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (ComponentSettings.DataGridViewCheckIfTheValueExists(dgv_SortColumns, "Keyword", keyword))
        {
            MessageBox.Show($"\"{keyword}\" already exists!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var index = dgv_SortColumns.Rows.Add();
        dgv_SortColumns.Rows[index].Cells["Keyword"].Value = keyword;
        txt_SortColumns.Text = string.Empty;
        txt_SortColumns.Focus();
    }

    private void Btn_SortColumnsRemove_Click(object? sender, EventArgs e)
    {
        GetTheIndexOfTheSelectedRow(dgv_SortColumns);
        // if (dgv_SortColumns.SelectedRows.Count > 0)
        // {
        //     dgv_SortColumns.Rows.RemoveAt(dgv_SortColumns.SelectedRows[0].Index);
        // }
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

    private void btn_RemoveColumnsAdd_Click(object sender, EventArgs e)
    {
        var keyword = txt_RemoveColumns.Text;
        var filter = cmb_RemoveColumnsFilterOption.Text;

        if (string.IsNullOrEmpty(keyword) || string.IsNullOrWhiteSpace(keyword))
        {
            MessageBox.Show("Keyword cannot be empty!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (ComponentSettings.DataGridViewCheckIfTheValueExists(dgv_RemoveColumns, "Keyword", keyword))
        {
            MessageBox.Show($"\"{keyword}\" already exists!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var index = dgv_RemoveColumns.Rows.Add();
        dgv_RemoveColumns.Rows[index].Cells["Filter"].Value = filter;
        dgv_RemoveColumns.Rows[index].Cells["Keyword"].Value = keyword;

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

    private void Btn_SortColumnsMoveToUp_Click(object? sender, EventArgs e)
    {
        MoveRowInDataGridViewToUp(dgv_SortColumns);
    }

    private void Btn_SortColumnsMoveToDown_Click(object? sender, EventArgs e)
    {
        MoveRowInDataGridViewToDown(dgv_SortColumns);
    }

    private void MoveRowInDataGridViewToUp(DataGridView dataGridView){
        if (dataGridView.SelectedRows.Count > 0)
        {
            var index = dataGridView.SelectedRows[0].Index;
            if (index == 0) { return; }
            var row = dataGridView.Rows[index];
            dataGridView.Rows.Remove(row);
            dataGridView.Rows.Insert(index - 1, row);
            dataGridView.ClearSelection();
            dataGridView.Rows[index - 1].Selected = true;
        }
    }

    private void MoveRowInDataGridViewToDown(DataGridView dataGridView){
        if (dataGridView.SelectedRows.Count > 0)
        {
            var index = dataGridView.SelectedRows[0].Index;
            if (index == dataGridView.Rows.Count - 1) { return; }
            var row = dataGridView.Rows[index];
            dataGridView.Rows.Remove(row);
            dataGridView.Rows.Insert(index + 1, row);
            dataGridView.ClearSelection();
            dataGridView.Rows[index + 1].Selected = true;
        }
    }

    private void GetTheIndexOfTheSelectedRow(DataGridView dataGridView){
        if (dataGridView.SelectedRows.Count > 0)
        {
            var index = dataGridView.SelectedRows[0].Index;
            MessageBox.Show(index.ToString());
        }
    }
}
