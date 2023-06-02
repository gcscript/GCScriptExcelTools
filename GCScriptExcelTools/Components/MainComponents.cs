using GCScriptExcelTools.Components;

namespace GCScriptExcelTools;
public partial class frm_Main
{
    public TableLayoutPanel Tlp_Main = new();
    public Panel Pnl_Tests = new();
    private void StartForm()
    {
        BackColor = ComponentsColor.DarkPrimary;
        Font = new Font("Consolas", 9F, FontStyle.Regular);
        MaximizeBox = false;
        MinimumSize = new Size(700, 600);
        Name = "frm_Main";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "GCScript Excel Tools";

        Tlp_Main_Component();




        Pnl_Tests.BackColor = Color.Blue;
        Pnl_Tests.Dock = DockStyle.Fill;
        Tlp_Main.Controls.Add(Pnl_Tests, 0, 0);
        Tlp_Main.SetColumnSpan(Pnl_Tests, 1);

    }

    private void Tlp_Main_Component()
    {
        Tlp_Main.Name = "Tlp_Main";
        Tlp_Main.BackColor = Color.Red;
        Tlp_Main.Visible = true;
        Tlp_Main.Enabled = true;
        Tlp_Main.Dock = DockStyle.Fill;
        Tlp_Main.ColumnCount = 1;
        Tlp_Main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        Tlp_Main.RowCount = 4;
        Tlp_Main.RowStyles.Add(new RowStyle(SizeType.Absolute, 200F));
        Tlp_Main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        Tlp_Main.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
        Tlp_Main.RowStyles.Add(new RowStyle(SizeType.Absolute, 15F));
        Controls.Add(Tlp_Main);
    }
}
