using GCScriptExcelTools.Components;

namespace GCScriptExcelTools;
public partial class frm_MainOld
{
    public DataGridView dgv_RemoveColumns = new();
    public DataGridView dgv_RenameColumns = new();

    public DataGridView dgv_SortColumns = new();
    public CheckBox chk_SortColumns = new();
    public TableLayoutPanel tlp_SortColumns = new();
    public TextBox txt_SortColumns = new();
    public Button btn_SortColumnsAdd = new();
    public Button btn_SortColumnsRemove = new();
    public Button btn_SortColumnsMoveToUp = new();
    public Button btn_SortColumnsMoveToDown = new();

    public void RemoveColumnsComponent()
    {
        ComponentSettings.DataGridViewDarkMode(dgv_RemoveColumns);
        ComponentSettings.CreateColumnDataGridView(dgv_RemoveColumns, "Filter", "Filter", ComponentSettings.EDataGridViewCellTemplate.TextBox, DataGridViewAutoSizeColumnMode.None);
        ComponentSettings.CreateColumnDataGridView(dgv_RemoveColumns, "Keyword", "Keyword", ComponentSettings.EDataGridViewCellTemplate.TextBox, DataGridViewAutoSizeColumnMode.Fill);
        pnl_RemoveColumns.Controls.Add(dgv_RemoveColumns);
    }

    public void RenameColumnsComponent()
    {

        ComponentSettings.DataGridViewDarkMode(dgv_RenameColumns);
        ComponentSettings.CreateColumnDataGridView(dgv_RenameColumns, "Filter", "Filter", ComponentSettings.EDataGridViewCellTemplate.TextBox, DataGridViewAutoSizeColumnMode.None);
        ComponentSettings.CreateColumnDataGridView(dgv_RenameColumns, "Find", "Find", ComponentSettings.EDataGridViewCellTemplate.TextBox, DataGridViewAutoSizeColumnMode.Fill);
        ComponentSettings.CreateColumnDataGridView(dgv_RenameColumns, "Replace", "Replace", ComponentSettings.EDataGridViewCellTemplate.TextBox, DataGridViewAutoSizeColumnMode.Fill);
        pnl_RenameColumns.Controls.Add(dgv_RenameColumns);
    }

    public void SortColumnsComponent()
    {
        chk_SortColumns.Text = "Sort Columns";
        chk_SortColumns.Name = "chk_SortColumns";
        chk_SortColumns.Dock = DockStyle.Top;
        chk_SortColumns.CheckedChanged += Chk_SortColumns_CheckedChanged;
        flp_Others.Controls.Add(chk_SortColumns);

        tlp_SortColumns.Name = "tlp_SortColumns";
        tlp_SortColumns.BackColor = ComponentsColor.DarkTertiary;
        tlp_SortColumns.Visible = false;
        tlp_SortColumns.Enabled = false;
        tlp_SortColumns.Dock = DockStyle.None;
        tlp_SortColumns.Size = new Size(flp_Others.Width - 25, 150);
        tlp_SortColumns.ColumnCount = 4;
        tlp_SortColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tlp_SortColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50F));
        tlp_SortColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50F));
        tlp_SortColumns.RowCount = 3;
        tlp_SortColumns.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
        tlp_SortColumns.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        tlp_SortColumns.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        flp_Others.Controls.Add(tlp_SortColumns);

        txt_SortColumns.BackColor = ComponentsColor.DarkPrimary;
        txt_SortColumns.ForeColor = ComponentsColor.LightPrimary;
        txt_SortColumns.BorderStyle = BorderStyle.FixedSingle;
        txt_SortColumns.Dock = DockStyle.Fill;
        txt_SortColumns.Name = "txt_SortColumns";
        txt_SortColumns.PlaceholderText = "Find what";
        tlp_SortColumns.Controls.Add(txt_SortColumns, 0, 0);
        tlp_SortColumns.SetColumnSpan(txt_SortColumns, 1);

        btn_SortColumnsAdd.BackColor = ComponentsColor.DarkPrimary;
        btn_SortColumnsAdd.ForeColor = ComponentsColor.LightPrimary;
        btn_SortColumnsAdd.FlatStyle = FlatStyle.Flat;
        btn_SortColumnsAdd.Dock = DockStyle.Fill;
        btn_SortColumnsAdd.Name = "btn_SortColumnsAdd";
        btn_SortColumnsAdd.Text = "+";
        btn_SortColumnsAdd.TextAlign = ContentAlignment.MiddleCenter;
        btn_SortColumnsAdd.Click += Btn_SortColumnsAdd_Click;
        tlp_SortColumns.Controls.Add(btn_SortColumnsAdd, 1, 0);
        tlp_SortColumns.SetColumnSpan(btn_SortColumnsAdd, 1);

        btn_SortColumnsRemove.BackColor = ComponentsColor.DarkPrimary;
        btn_SortColumnsRemove.ForeColor = ComponentsColor.LightPrimary;
        btn_SortColumnsRemove.FlatStyle = FlatStyle.Flat;
        btn_SortColumnsRemove.Dock = DockStyle.Fill;
        btn_SortColumnsRemove.Name = "btn_SortColumnsRemove";
        btn_SortColumnsRemove.Text = "-";
        btn_SortColumnsRemove.TextAlign = ContentAlignment.MiddleCenter;
        btn_SortColumnsRemove.Click += Btn_SortColumnsRemove_Click;
        tlp_SortColumns.Controls.Add(btn_SortColumnsRemove, 2, 0);
        tlp_SortColumns.SetColumnSpan(btn_SortColumnsRemove, 1);

        btn_SortColumnsMoveToUp.BackColor = ComponentsColor.DarkPrimary;
        btn_SortColumnsMoveToUp.ForeColor = ComponentsColor.LightPrimary;
        btn_SortColumnsMoveToUp.FlatStyle = FlatStyle.Flat;
        btn_SortColumnsMoveToUp.Dock = DockStyle.Fill;
        btn_SortColumnsMoveToUp.Name = "btn_SortColumnsMoveToUp";
        btn_SortColumnsMoveToUp.Text = "↑";
        btn_SortColumnsMoveToUp.TextAlign = ContentAlignment.MiddleCenter;
        btn_SortColumnsMoveToUp.Click += Btn_SortColumnsMoveToUp_Click;
        tlp_SortColumns.Controls.Add(btn_SortColumnsMoveToUp, 2, 1);
        tlp_SortColumns.SetColumnSpan(btn_SortColumnsMoveToUp, 1);

        btn_SortColumnsMoveToDown.BackColor = ComponentsColor.DarkPrimary;
        btn_SortColumnsMoveToDown.ForeColor = ComponentsColor.LightPrimary;
        btn_SortColumnsMoveToDown.FlatStyle = FlatStyle.Flat;
        btn_SortColumnsMoveToDown.Dock = DockStyle.Fill;
        btn_SortColumnsMoveToDown.Name = "btn_SortColumnsMoveToDown";
        btn_SortColumnsMoveToDown.Text = "↓";
        btn_SortColumnsMoveToDown.TextAlign = ContentAlignment.MiddleCenter;
        btn_SortColumnsMoveToDown.Click += Btn_SortColumnsMoveToDown_Click;
        tlp_SortColumns.Controls.Add(btn_SortColumnsMoveToDown, 2, 2);
        tlp_SortColumns.SetColumnSpan(btn_SortColumnsMoveToDown, 1);

        ComponentSettings.DataGridViewDarkMode(dgv_SortColumns);
        ComponentSettings.CreateColumnDataGridView(dgv_SortColumns, "Keyword", "Order", ComponentSettings.EDataGridViewCellTemplate.TextBox, DataGridViewAutoSizeColumnMode.Fill);
        tlp_SortColumns.Controls.Add(dgv_SortColumns, 0, 1);
        tlp_SortColumns.SetRowSpan(dgv_SortColumns, 2);
        tlp_SortColumns.SetColumnSpan(dgv_SortColumns, 2);
    }
}
