namespace GCScriptExcelTools
{
    partial class frm_Main
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm_Main));
            btn_Add = new Button();
            lst_FilesPath = new ListBox();
            label1 = new Label();
            tableLayoutPanel1 = new TableLayoutPanel();
            btn_RemoveFile = new Button();
            panel1 = new Panel();
            btn_Start = new Button();
            flp_Remove = new FlowLayoutPanel();
            chk_RemoveInvisibleWorksheets = new CheckBox();
            chk_RemoveEmptyWorksheets = new CheckBox();
            chk_RemoveEmptyRows = new CheckBox();
            chk_RemoveEmptyColumns = new CheckBox();
            chk_RemoveHiddenRows = new CheckBox();
            chk_RemoveFontColor = new CheckBox();
            chk_RemoveBackgroundColor = new CheckBox();
            chk_RemoveFormatting = new CheckBox();
            chk_RemoveColumns = new CheckBox();
            tlp_RemoveColumns = new TableLayoutPanel();
            cmb_RemoveColumnsFilterOption = new ComboBox();
            btn_RemoveColumnsAdd = new Button();
            btn_RemoveColumnsRemove = new Button();
            txt_RemoveColumns = new TextBox();
            dgv_RemoveColumns = new DataGridView();
            flp_Apply = new FlowLayoutPanel();
            chk_SortWorksheets = new CheckBox();
            tlp_SortWorksheets = new TableLayoutPanel();
            cmb_SortWorksheets = new ComboBox();
            chk_Font = new CheckBox();
            tlp_Font = new TableLayoutPanel();
            nud_FontSize = new NumericUpDown();
            cmb_FontName = new ComboBox();
            chk_CellAlignments = new CheckBox();
            tlp_CellAlignments = new TableLayoutPanel();
            cmb_HorizontalCellAlignments = new ComboBox();
            cmb_VerticalCellAlignments = new ComboBox();
            lbl_HorizontalCellAlignments = new Label();
            lbl_VerticalCellAlignments = new Label();
            chk_RowHeight = new CheckBox();
            tlp_RowHeight = new TableLayoutPanel();
            label8 = new Label();
            chk_RowHeightAuto = new CheckBox();
            nud_RowHeight = new NumericUpDown();
            nud_RowMaxHeight = new NumericUpDown();
            chk_ColumnWidth = new CheckBox();
            tlp_ColumnWidth = new TableLayoutPanel();
            label9 = new Label();
            chk_ColumnWidthAuto = new CheckBox();
            nud_ColumnWidth = new NumericUpDown();
            nud_ColumnMaxWidth = new NumericUpDown();
            chk_Zoom = new CheckBox();
            tlp_Zoom = new TableLayoutPanel();
            nud_Zoom = new NumericUpDown();
            label10 = new Label();
            tableLayoutPanel3 = new TableLayoutPanel();
            tlp_Main = new TableLayoutPanel();
            tableLayoutPanel2 = new TableLayoutPanel();
            tableLayoutPanel5 = new TableLayoutPanel();
            cmb_Preset = new ComboBox();
            lbl_Preset = new Label();
            btn_SavePreset = new Button();
            btn_RemovePreset = new Button();
            pnl_Options = new Panel();
            pnl_Apply = new Panel();
            pnl_Remove = new Panel();
            pnl_Others = new Panel();
            flp_Others = new FlowLayoutPanel();
            chk_GetLastRealEmptyRow = new CheckBox();
            chk_GetLastRealEmptyColumn = new CheckBox();
            chk_FindHeader = new CheckBox();
            tlp_FindHeader = new TableLayoutPanel();
            cmb_FindHeaderFilterOption = new ComboBox();
            txt_FindHeader = new TextBox();
            btn_FindHeaderAdd = new Button();
            btn_FindHeaderRemove = new Button();
            pnl_FindHeader = new Panel();
            lst_FindHeader = new ListBox();
            tableLayoutPanel4 = new TableLayoutPanel();
            btn_Apply = new Button();
            btn_Remove = new Button();
            btn_Others = new Button();
            lbl_Credits = new Label();
            tableLayoutPanel1.SuspendLayout();
            panel1.SuspendLayout();
            flp_Remove.SuspendLayout();
            tlp_RemoveColumns.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgv_RemoveColumns).BeginInit();
            flp_Apply.SuspendLayout();
            tlp_SortWorksheets.SuspendLayout();
            tlp_Font.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)nud_FontSize).BeginInit();
            tlp_CellAlignments.SuspendLayout();
            tlp_RowHeight.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)nud_RowHeight).BeginInit();
            ((System.ComponentModel.ISupportInitialize)nud_RowMaxHeight).BeginInit();
            tlp_ColumnWidth.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)nud_ColumnWidth).BeginInit();
            ((System.ComponentModel.ISupportInitialize)nud_ColumnMaxWidth).BeginInit();
            tlp_Zoom.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)nud_Zoom).BeginInit();
            tableLayoutPanel3.SuspendLayout();
            tlp_Main.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            tableLayoutPanel5.SuspendLayout();
            pnl_Options.SuspendLayout();
            pnl_Apply.SuspendLayout();
            pnl_Remove.SuspendLayout();
            pnl_Others.SuspendLayout();
            flp_Others.SuspendLayout();
            tlp_FindHeader.SuspendLayout();
            pnl_FindHeader.SuspendLayout();
            tableLayoutPanel4.SuspendLayout();
            SuspendLayout();
            // 
            // btn_Add
            // 
            btn_Add.BackColor = Color.FromArgb(25, 135, 84);
            btn_Add.Dock = DockStyle.Fill;
            btn_Add.FlatStyle = FlatStyle.Flat;
            btn_Add.ForeColor = Color.White;
            btn_Add.Location = new Point(586, 18);
            btn_Add.Name = "btn_Add";
            btn_Add.Size = new Size(95, 39);
            btn_Add.TabIndex = 2;
            btn_Add.Text = "Add Files";
            btn_Add.UseVisualStyleBackColor = false;
            btn_Add.Click += btn_Add_Click;
            // 
            // lst_FilesPath
            // 
            lst_FilesPath.BackColor = Color.FromArgb(40, 40, 40);
            lst_FilesPath.BorderStyle = BorderStyle.None;
            lst_FilesPath.Dock = DockStyle.Fill;
            lst_FilesPath.ForeColor = Color.FromArgb(224, 224, 224);
            lst_FilesPath.FormattingEnabled = true;
            lst_FilesPath.ItemHeight = 14;
            lst_FilesPath.Location = new Point(0, 0);
            lst_FilesPath.Name = "lst_FilesPath";
            lst_FilesPath.Size = new Size(577, 85);
            lst_FilesPath.TabIndex = 3;
            lst_FilesPath.SelectedIndexChanged += lst_FilesPath_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.BackColor = Color.FromArgb(20, 20, 20);
            tableLayoutPanel1.SetColumnSpan(label1, 2);
            label1.Dock = DockStyle.Fill;
            label1.Font = new Font("Consolas", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            label1.ForeColor = Color.White;
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(577, 15);
            label1.TabIndex = 4;
            label1.Text = "File List";
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.BackColor = Color.FromArgb(20, 20, 20);
            tableLayoutPanel1.ColumnCount = 3;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel1.Controls.Add(btn_Add, 2, 1);
            tableLayoutPanel1.Controls.Add(btn_RemoveFile, 2, 2);
            tableLayoutPanel1.Controls.Add(label1, 0, 0);
            tableLayoutPanel1.Controls.Add(panel1, 0, 1);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.ForeColor = Color.White;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Margin = new Padding(0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 3;
            tableLayoutPanel1.RowStyles.Add(new RowStyle());
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Size = new Size(684, 106);
            tableLayoutPanel1.TabIndex = 5;
            // 
            // btn_RemoveFile
            // 
            btn_RemoveFile.BackColor = Color.FromArgb(220, 53, 69);
            btn_RemoveFile.Dock = DockStyle.Fill;
            btn_RemoveFile.FlatStyle = FlatStyle.Flat;
            btn_RemoveFile.ForeColor = Color.White;
            btn_RemoveFile.Location = new Point(586, 63);
            btn_RemoveFile.Name = "btn_RemoveFile";
            btn_RemoveFile.Size = new Size(95, 40);
            btn_RemoveFile.TabIndex = 5;
            btn_RemoveFile.Text = "Remove Selected";
            btn_RemoveFile.UseVisualStyleBackColor = false;
            btn_RemoveFile.Click += btn_RemoveFile_Click;
            // 
            // panel1
            // 
            panel1.BackColor = Color.FromArgb(20, 20, 20);
            tableLayoutPanel1.SetColumnSpan(panel1, 2);
            panel1.Controls.Add(lst_FilesPath);
            panel1.Dock = DockStyle.Fill;
            panel1.ForeColor = Color.White;
            panel1.Location = new Point(3, 18);
            panel1.Name = "panel1";
            tableLayoutPanel1.SetRowSpan(panel1, 2);
            panel1.Size = new Size(577, 85);
            panel1.TabIndex = 6;
            // 
            // btn_Start
            // 
            btn_Start.BackColor = Color.FromArgb(25, 135, 84);
            tableLayoutPanel3.SetColumnSpan(btn_Start, 2);
            btn_Start.Dock = DockStyle.Fill;
            btn_Start.Enabled = false;
            btn_Start.FlatStyle = FlatStyle.Flat;
            btn_Start.ForeColor = Color.White;
            btn_Start.Location = new Point(3, 3);
            btn_Start.Name = "btn_Start";
            btn_Start.Size = new Size(678, 44);
            btn_Start.TabIndex = 7;
            btn_Start.Text = "Start";
            btn_Start.UseVisualStyleBackColor = false;
            btn_Start.Click += btn_Start_Click;
            // 
            // flp_Remove
            // 
            flp_Remove.AutoScroll = true;
            flp_Remove.AutoSize = true;
            flp_Remove.BackColor = Color.FromArgb(20, 20, 20);
            flp_Remove.BorderStyle = BorderStyle.FixedSingle;
            flp_Remove.Controls.Add(chk_RemoveInvisibleWorksheets);
            flp_Remove.Controls.Add(chk_RemoveEmptyWorksheets);
            flp_Remove.Controls.Add(chk_RemoveEmptyRows);
            flp_Remove.Controls.Add(chk_RemoveEmptyColumns);
            flp_Remove.Controls.Add(chk_RemoveHiddenRows);
            flp_Remove.Controls.Add(chk_RemoveFontColor);
            flp_Remove.Controls.Add(chk_RemoveBackgroundColor);
            flp_Remove.Controls.Add(chk_RemoveFormatting);
            flp_Remove.Controls.Add(chk_RemoveColumns);
            flp_Remove.Controls.Add(tlp_RemoveColumns);
            flp_Remove.Dock = DockStyle.Fill;
            flp_Remove.FlowDirection = FlowDirection.TopDown;
            flp_Remove.ForeColor = Color.White;
            flp_Remove.Location = new Point(0, 0);
            flp_Remove.Name = "flp_Remove";
            flp_Remove.Size = new Size(672, 308);
            flp_Remove.TabIndex = 10;
            flp_Remove.WrapContents = false;
            // 
            // chk_RemoveInvisibleWorksheets
            // 
            chk_RemoveInvisibleWorksheets.AutoSize = true;
            chk_RemoveInvisibleWorksheets.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveInvisibleWorksheets.Checked = true;
            chk_RemoveInvisibleWorksheets.CheckState = CheckState.Checked;
            chk_RemoveInvisibleWorksheets.ForeColor = Color.White;
            chk_RemoveInvisibleWorksheets.Location = new Point(3, 3);
            chk_RemoveInvisibleWorksheets.Name = "chk_RemoveInvisibleWorksheets";
            chk_RemoveInvisibleWorksheets.Size = new Size(166, 18);
            chk_RemoveInvisibleWorksheets.TabIndex = 4;
            chk_RemoveInvisibleWorksheets.Text = "Invisible Worksheets";
            chk_RemoveInvisibleWorksheets.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveEmptyWorksheets
            // 
            chk_RemoveEmptyWorksheets.AutoSize = true;
            chk_RemoveEmptyWorksheets.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveEmptyWorksheets.Checked = true;
            chk_RemoveEmptyWorksheets.CheckState = CheckState.Checked;
            chk_RemoveEmptyWorksheets.ForeColor = Color.White;
            chk_RemoveEmptyWorksheets.Location = new Point(3, 27);
            chk_RemoveEmptyWorksheets.Name = "chk_RemoveEmptyWorksheets";
            chk_RemoveEmptyWorksheets.Size = new Size(138, 18);
            chk_RemoveEmptyWorksheets.TabIndex = 0;
            chk_RemoveEmptyWorksheets.Text = "Empty Worksheets";
            chk_RemoveEmptyWorksheets.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveEmptyRows
            // 
            chk_RemoveEmptyRows.AutoSize = true;
            chk_RemoveEmptyRows.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveEmptyRows.Checked = true;
            chk_RemoveEmptyRows.CheckState = CheckState.Checked;
            chk_RemoveEmptyRows.ForeColor = Color.White;
            chk_RemoveEmptyRows.Location = new Point(3, 51);
            chk_RemoveEmptyRows.Name = "chk_RemoveEmptyRows";
            chk_RemoveEmptyRows.Size = new Size(96, 18);
            chk_RemoveEmptyRows.TabIndex = 1;
            chk_RemoveEmptyRows.Text = "Empty Rows";
            chk_RemoveEmptyRows.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveEmptyColumns
            // 
            chk_RemoveEmptyColumns.AutoSize = true;
            chk_RemoveEmptyColumns.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveEmptyColumns.Checked = true;
            chk_RemoveEmptyColumns.CheckState = CheckState.Checked;
            chk_RemoveEmptyColumns.ForeColor = Color.White;
            chk_RemoveEmptyColumns.Location = new Point(3, 75);
            chk_RemoveEmptyColumns.Name = "chk_RemoveEmptyColumns";
            chk_RemoveEmptyColumns.Size = new Size(117, 18);
            chk_RemoveEmptyColumns.TabIndex = 1;
            chk_RemoveEmptyColumns.Text = "Empty Columns";
            chk_RemoveEmptyColumns.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveHiddenRows
            // 
            chk_RemoveHiddenRows.AutoSize = true;
            chk_RemoveHiddenRows.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveHiddenRows.Checked = true;
            chk_RemoveHiddenRows.CheckState = CheckState.Checked;
            chk_RemoveHiddenRows.ForeColor = Color.White;
            chk_RemoveHiddenRows.Location = new Point(3, 99);
            chk_RemoveHiddenRows.Name = "chk_RemoveHiddenRows";
            chk_RemoveHiddenRows.Size = new Size(215, 18);
            chk_RemoveHiddenRows.TabIndex = 21;
            chk_RemoveHiddenRows.Text = "Hidden Rows (Also Filtered)";
            chk_RemoveHiddenRows.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveFontColor
            // 
            chk_RemoveFontColor.AutoSize = true;
            chk_RemoveFontColor.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveFontColor.Checked = true;
            chk_RemoveFontColor.CheckState = CheckState.Checked;
            chk_RemoveFontColor.ForeColor = Color.White;
            chk_RemoveFontColor.Location = new Point(3, 123);
            chk_RemoveFontColor.Name = "chk_RemoveFontColor";
            chk_RemoveFontColor.Size = new Size(96, 18);
            chk_RemoveFontColor.TabIndex = 3;
            chk_RemoveFontColor.Text = "Font Color";
            chk_RemoveFontColor.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveBackgroundColor
            // 
            chk_RemoveBackgroundColor.AutoSize = true;
            chk_RemoveBackgroundColor.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveBackgroundColor.Checked = true;
            chk_RemoveBackgroundColor.CheckState = CheckState.Checked;
            chk_RemoveBackgroundColor.ForeColor = Color.White;
            chk_RemoveBackgroundColor.Location = new Point(3, 147);
            chk_RemoveBackgroundColor.Name = "chk_RemoveBackgroundColor";
            chk_RemoveBackgroundColor.Size = new Size(138, 18);
            chk_RemoveBackgroundColor.TabIndex = 2;
            chk_RemoveBackgroundColor.Text = "Background Color";
            chk_RemoveBackgroundColor.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveFormatting
            // 
            chk_RemoveFormatting.AutoSize = true;
            chk_RemoveFormatting.BackColor = Color.FromArgb(20, 20, 20);
            chk_RemoveFormatting.ForeColor = Color.White;
            chk_RemoveFormatting.Location = new Point(3, 171);
            chk_RemoveFormatting.Name = "chk_RemoveFormatting";
            chk_RemoveFormatting.Size = new Size(96, 18);
            chk_RemoveFormatting.TabIndex = 1;
            chk_RemoveFormatting.Text = "Formatting";
            chk_RemoveFormatting.UseVisualStyleBackColor = false;
            // 
            // chk_RemoveColumns
            // 
            chk_RemoveColumns.AutoSize = true;
            chk_RemoveColumns.Checked = true;
            chk_RemoveColumns.CheckState = CheckState.Checked;
            chk_RemoveColumns.Location = new Point(3, 195);
            chk_RemoveColumns.Name = "chk_RemoveColumns";
            chk_RemoveColumns.Size = new Size(75, 18);
            chk_RemoveColumns.TabIndex = 20;
            chk_RemoveColumns.Text = "Columns";
            chk_RemoveColumns.UseVisualStyleBackColor = true;
            chk_RemoveColumns.CheckedChanged += chk_RemoveColumns_CheckedChanged;
            // 
            // tlp_RemoveColumns
            // 
            tlp_RemoveColumns.BackColor = Color.FromArgb(40, 40, 40);
            tlp_RemoveColumns.ColumnCount = 10;
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_RemoveColumns.Controls.Add(cmb_RemoveColumnsFilterOption, 0, 0);
            tlp_RemoveColumns.Controls.Add(btn_RemoveColumnsAdd, 8, 0);
            tlp_RemoveColumns.Controls.Add(btn_RemoveColumnsRemove, 9, 0);
            tlp_RemoveColumns.Controls.Add(txt_RemoveColumns, 2, 0);
            tlp_RemoveColumns.Controls.Add(dgv_RemoveColumns, 0, 1);
            tlp_RemoveColumns.Location = new Point(3, 219);
            tlp_RemoveColumns.Name = "tlp_RemoveColumns";
            tlp_RemoveColumns.RowCount = 5;
            tlp_RemoveColumns.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_RemoveColumns.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_RemoveColumns.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_RemoveColumns.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_RemoveColumns.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_RemoveColumns.Size = new Size(645, 140);
            tlp_RemoveColumns.TabIndex = 19;
            // 
            // cmb_RemoveColumnsFilterOption
            // 
            cmb_RemoveColumnsFilterOption.BackColor = Color.FromArgb(20, 20, 20);
            tlp_RemoveColumns.SetColumnSpan(cmb_RemoveColumnsFilterOption, 2);
            cmb_RemoveColumnsFilterOption.Dock = DockStyle.Fill;
            cmb_RemoveColumnsFilterOption.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_RemoveColumnsFilterOption.FlatStyle = FlatStyle.Flat;
            cmb_RemoveColumnsFilterOption.ForeColor = Color.FromArgb(224, 224, 224);
            cmb_RemoveColumnsFilterOption.FormattingEnabled = true;
            cmb_RemoveColumnsFilterOption.Items.AddRange(new object[] { "Equals", "Contains", "Starts With", "Ends With" });
            cmb_RemoveColumnsFilterOption.Location = new Point(3, 3);
            cmb_RemoveColumnsFilterOption.Name = "cmb_RemoveColumnsFilterOption";
            cmb_RemoveColumnsFilterOption.Size = new Size(122, 22);
            cmb_RemoveColumnsFilterOption.TabIndex = 5;
            // 
            // btn_RemoveColumnsAdd
            // 
            btn_RemoveColumnsAdd.BackColor = Color.FromArgb(40, 40, 40);
            btn_RemoveColumnsAdd.Dock = DockStyle.Fill;
            btn_RemoveColumnsAdd.FlatStyle = FlatStyle.Flat;
            btn_RemoveColumnsAdd.Location = new Point(515, 3);
            btn_RemoveColumnsAdd.Name = "btn_RemoveColumnsAdd";
            btn_RemoveColumnsAdd.Size = new Size(58, 22);
            btn_RemoveColumnsAdd.TabIndex = 0;
            btn_RemoveColumnsAdd.Text = "+";
            btn_RemoveColumnsAdd.UseVisualStyleBackColor = false;
            btn_RemoveColumnsAdd.Click += btn_RemoveColumnsAdd_Click;
            // 
            // btn_RemoveColumnsRemove
            // 
            btn_RemoveColumnsRemove.BackColor = Color.FromArgb(40, 40, 40);
            btn_RemoveColumnsRemove.Dock = DockStyle.Fill;
            btn_RemoveColumnsRemove.FlatStyle = FlatStyle.Flat;
            btn_RemoveColumnsRemove.Location = new Point(579, 3);
            btn_RemoveColumnsRemove.Name = "btn_RemoveColumnsRemove";
            btn_RemoveColumnsRemove.Size = new Size(63, 22);
            btn_RemoveColumnsRemove.TabIndex = 1;
            btn_RemoveColumnsRemove.Text = "-";
            btn_RemoveColumnsRemove.UseVisualStyleBackColor = false;
            btn_RemoveColumnsRemove.Click += btn_RemoveColumnsRemove_Click;
            // 
            // txt_RemoveColumns
            // 
            txt_RemoveColumns.BackColor = Color.FromArgb(20, 20, 20);
            txt_RemoveColumns.BorderStyle = BorderStyle.FixedSingle;
            tlp_RemoveColumns.SetColumnSpan(txt_RemoveColumns, 6);
            txt_RemoveColumns.Dock = DockStyle.Fill;
            txt_RemoveColumns.ForeColor = Color.FromArgb(224, 224, 224);
            txt_RemoveColumns.Location = new Point(131, 3);
            txt_RemoveColumns.Name = "txt_RemoveColumns";
            txt_RemoveColumns.PlaceholderText = "Enter your text here";
            txt_RemoveColumns.Size = new Size(378, 22);
            txt_RemoveColumns.TabIndex = 3;
            // 
            // dgv_RemoveColumns
            // 
            dgv_RemoveColumns.AllowUserToAddRows = false;
            dgv_RemoveColumns.AllowUserToDeleteRows = false;
            dgv_RemoveColumns.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            tlp_RemoveColumns.SetColumnSpan(dgv_RemoveColumns, 10);
            dgv_RemoveColumns.Dock = DockStyle.Fill;
            dgv_RemoveColumns.Location = new Point(3, 31);
            dgv_RemoveColumns.Name = "dgv_RemoveColumns";
            dgv_RemoveColumns.ReadOnly = true;
            tlp_RemoveColumns.SetRowSpan(dgv_RemoveColumns, 4);
            dgv_RemoveColumns.RowTemplate.Height = 25;
            dgv_RemoveColumns.Size = new Size(639, 106);
            dgv_RemoveColumns.TabIndex = 6;
            // 
            // flp_Apply
            // 
            flp_Apply.AutoScroll = true;
            flp_Apply.AutoSize = true;
            flp_Apply.BackColor = Color.FromArgb(20, 20, 20);
            flp_Apply.BorderStyle = BorderStyle.FixedSingle;
            flp_Apply.Controls.Add(chk_SortWorksheets);
            flp_Apply.Controls.Add(tlp_SortWorksheets);
            flp_Apply.Controls.Add(chk_Font);
            flp_Apply.Controls.Add(tlp_Font);
            flp_Apply.Controls.Add(chk_CellAlignments);
            flp_Apply.Controls.Add(tlp_CellAlignments);
            flp_Apply.Controls.Add(chk_RowHeight);
            flp_Apply.Controls.Add(tlp_RowHeight);
            flp_Apply.Controls.Add(chk_ColumnWidth);
            flp_Apply.Controls.Add(tlp_ColumnWidth);
            flp_Apply.Controls.Add(chk_Zoom);
            flp_Apply.Controls.Add(tlp_Zoom);
            flp_Apply.Dock = DockStyle.Fill;
            flp_Apply.FlowDirection = FlowDirection.TopDown;
            flp_Apply.ForeColor = Color.White;
            flp_Apply.Location = new Point(0, 0);
            flp_Apply.Name = "flp_Apply";
            flp_Apply.Size = new Size(672, 308);
            flp_Apply.TabIndex = 9;
            flp_Apply.WrapContents = false;
            // 
            // chk_SortWorksheets
            // 
            chk_SortWorksheets.AutoSize = true;
            chk_SortWorksheets.BackColor = Color.FromArgb(20, 20, 20);
            chk_SortWorksheets.Checked = true;
            chk_SortWorksheets.CheckState = CheckState.Checked;
            chk_SortWorksheets.ForeColor = Color.White;
            chk_SortWorksheets.Location = new Point(3, 3);
            chk_SortWorksheets.Name = "chk_SortWorksheets";
            chk_SortWorksheets.Size = new Size(131, 18);
            chk_SortWorksheets.TabIndex = 0;
            chk_SortWorksheets.Text = "Sort Worksheets";
            chk_SortWorksheets.UseVisualStyleBackColor = false;
            chk_SortWorksheets.CheckedChanged += chk_SortWorksheets_CheckedChanged;
            // 
            // tlp_SortWorksheets
            // 
            tlp_SortWorksheets.BackColor = Color.FromArgb(40, 40, 40);
            tlp_SortWorksheets.ColumnCount = 2;
            tlp_SortWorksheets.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlp_SortWorksheets.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56F));
            tlp_SortWorksheets.Controls.Add(cmb_SortWorksheets, 0, 0);
            tlp_SortWorksheets.Location = new Point(3, 27);
            tlp_SortWorksheets.Name = "tlp_SortWorksheets";
            tlp_SortWorksheets.RowCount = 1;
            tlp_SortWorksheets.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_SortWorksheets.Size = new Size(645, 28);
            tlp_SortWorksheets.TabIndex = 16;
            // 
            // cmb_SortWorksheets
            // 
            cmb_SortWorksheets.BackColor = Color.FromArgb(20, 20, 20);
            tlp_SortWorksheets.SetColumnSpan(cmb_SortWorksheets, 2);
            cmb_SortWorksheets.Dock = DockStyle.Fill;
            cmb_SortWorksheets.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_SortWorksheets.FlatStyle = FlatStyle.Flat;
            cmb_SortWorksheets.ForeColor = SystemColors.HighlightText;
            cmb_SortWorksheets.FormattingEnabled = true;
            cmb_SortWorksheets.Items.AddRange(new object[] { "Ascending", "Descending" });
            cmb_SortWorksheets.Location = new Point(3, 3);
            cmb_SortWorksheets.Name = "cmb_SortWorksheets";
            cmb_SortWorksheets.Size = new Size(639, 22);
            cmb_SortWorksheets.TabIndex = 8;
            // 
            // chk_Font
            // 
            chk_Font.AutoSize = true;
            chk_Font.BackColor = Color.FromArgb(20, 20, 20);
            chk_Font.Checked = true;
            chk_Font.CheckState = CheckState.Checked;
            chk_Font.ForeColor = Color.White;
            chk_Font.Location = new Point(3, 61);
            chk_Font.Name = "chk_Font";
            chk_Font.Size = new Size(54, 18);
            chk_Font.TabIndex = 4;
            chk_Font.Text = "Font";
            chk_Font.UseVisualStyleBackColor = false;
            chk_Font.CheckedChanged += chk_FontSettings_CheckedChanged;
            // 
            // tlp_Font
            // 
            tlp_Font.BackColor = Color.FromArgb(40, 40, 40);
            tlp_Font.ColumnCount = 2;
            tlp_Font.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlp_Font.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56F));
            tlp_Font.Controls.Add(nud_FontSize, 1, 0);
            tlp_Font.Controls.Add(cmb_FontName, 0, 0);
            tlp_Font.Location = new Point(3, 85);
            tlp_Font.Name = "tlp_Font";
            tlp_Font.RowCount = 1;
            tlp_Font.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_Font.Size = new Size(645, 28);
            tlp_Font.TabIndex = 16;
            // 
            // nud_FontSize
            // 
            nud_FontSize.BackColor = Color.FromArgb(20, 20, 20);
            nud_FontSize.Dock = DockStyle.Fill;
            nud_FontSize.ForeColor = Color.White;
            nud_FontSize.Location = new Point(592, 3);
            nud_FontSize.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
            nud_FontSize.Minimum = new decimal(new int[] { 8, 0, 0, 0 });
            nud_FontSize.Name = "nud_FontSize";
            nud_FontSize.Size = new Size(50, 22);
            nud_FontSize.TabIndex = 6;
            nud_FontSize.TextAlign = HorizontalAlignment.Center;
            nud_FontSize.Value = new decimal(new int[] { 10, 0, 0, 0 });
            // 
            // cmb_FontName
            // 
            cmb_FontName.BackColor = Color.FromArgb(20, 20, 20);
            cmb_FontName.Dock = DockStyle.Fill;
            cmb_FontName.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_FontName.FlatStyle = FlatStyle.Flat;
            cmb_FontName.ForeColor = SystemColors.HighlightText;
            cmb_FontName.FormattingEnabled = true;
            cmb_FontName.Items.AddRange(new object[] { "Consolas" });
            cmb_FontName.Location = new Point(3, 3);
            cmb_FontName.Name = "cmb_FontName";
            cmb_FontName.Size = new Size(583, 22);
            cmb_FontName.TabIndex = 5;
            // 
            // chk_CellAlignments
            // 
            chk_CellAlignments.AutoSize = true;
            chk_CellAlignments.BackColor = Color.FromArgb(20, 20, 20);
            chk_CellAlignments.Checked = true;
            chk_CellAlignments.CheckState = CheckState.Checked;
            chk_CellAlignments.ForeColor = Color.White;
            chk_CellAlignments.Location = new Point(3, 119);
            chk_CellAlignments.Name = "chk_CellAlignments";
            chk_CellAlignments.Size = new Size(131, 18);
            chk_CellAlignments.TabIndex = 10;
            chk_CellAlignments.Text = "Cell Alignments";
            chk_CellAlignments.UseVisualStyleBackColor = false;
            chk_CellAlignments.CheckedChanged += chk_CellAlignments_CheckedChanged;
            // 
            // tlp_CellAlignments
            // 
            tlp_CellAlignments.BackColor = Color.FromArgb(40, 40, 40);
            tlp_CellAlignments.ColumnCount = 2;
            tlp_CellAlignments.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_CellAlignments.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_CellAlignments.Controls.Add(cmb_HorizontalCellAlignments, 1, 1);
            tlp_CellAlignments.Controls.Add(cmb_VerticalCellAlignments, 0, 1);
            tlp_CellAlignments.Controls.Add(lbl_HorizontalCellAlignments, 1, 0);
            tlp_CellAlignments.Controls.Add(lbl_VerticalCellAlignments, 0, 0);
            tlp_CellAlignments.Location = new Point(3, 143);
            tlp_CellAlignments.Name = "tlp_CellAlignments";
            tlp_CellAlignments.RowCount = 2;
            tlp_CellAlignments.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_CellAlignments.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_CellAlignments.Size = new Size(645, 44);
            tlp_CellAlignments.TabIndex = 16;
            // 
            // cmb_HorizontalCellAlignments
            // 
            cmb_HorizontalCellAlignments.BackColor = Color.FromArgb(20, 20, 20);
            cmb_HorizontalCellAlignments.Dock = DockStyle.Fill;
            cmb_HorizontalCellAlignments.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_HorizontalCellAlignments.FlatStyle = FlatStyle.Flat;
            cmb_HorizontalCellAlignments.ForeColor = SystemColors.HighlightText;
            cmb_HorizontalCellAlignments.FormattingEnabled = true;
            cmb_HorizontalCellAlignments.Items.AddRange(new object[] { "Left", "Center", "Right" });
            cmb_HorizontalCellAlignments.Location = new Point(325, 19);
            cmb_HorizontalCellAlignments.Name = "cmb_HorizontalCellAlignments";
            cmb_HorizontalCellAlignments.Size = new Size(317, 22);
            cmb_HorizontalCellAlignments.TabIndex = 6;
            // 
            // cmb_VerticalCellAlignments
            // 
            cmb_VerticalCellAlignments.BackColor = Color.FromArgb(20, 20, 20);
            cmb_VerticalCellAlignments.Dock = DockStyle.Fill;
            cmb_VerticalCellAlignments.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_VerticalCellAlignments.FlatStyle = FlatStyle.Flat;
            cmb_VerticalCellAlignments.ForeColor = SystemColors.HighlightText;
            cmb_VerticalCellAlignments.FormattingEnabled = true;
            cmb_VerticalCellAlignments.Items.AddRange(new object[] { "Top", "Center", "Bottom" });
            cmb_VerticalCellAlignments.Location = new Point(3, 19);
            cmb_VerticalCellAlignments.Name = "cmb_VerticalCellAlignments";
            cmb_VerticalCellAlignments.Size = new Size(316, 22);
            cmb_VerticalCellAlignments.TabIndex = 8;
            // 
            // lbl_HorizontalCellAlignments
            // 
            lbl_HorizontalCellAlignments.BackColor = Color.FromArgb(40, 40, 40);
            lbl_HorizontalCellAlignments.Dock = DockStyle.Fill;
            lbl_HorizontalCellAlignments.ForeColor = Color.White;
            lbl_HorizontalCellAlignments.Location = new Point(325, 0);
            lbl_HorizontalCellAlignments.Name = "lbl_HorizontalCellAlignments";
            lbl_HorizontalCellAlignments.Size = new Size(317, 16);
            lbl_HorizontalCellAlignments.TabIndex = 9;
            lbl_HorizontalCellAlignments.Text = "Horizontal";
            lbl_HorizontalCellAlignments.TextAlign = ContentAlignment.BottomCenter;
            // 
            // lbl_VerticalCellAlignments
            // 
            lbl_VerticalCellAlignments.BackColor = Color.FromArgb(40, 40, 40);
            lbl_VerticalCellAlignments.Dock = DockStyle.Fill;
            lbl_VerticalCellAlignments.ForeColor = Color.White;
            lbl_VerticalCellAlignments.Location = new Point(0, 0);
            lbl_VerticalCellAlignments.Margin = new Padding(0);
            lbl_VerticalCellAlignments.Name = "lbl_VerticalCellAlignments";
            lbl_VerticalCellAlignments.Size = new Size(322, 16);
            lbl_VerticalCellAlignments.TabIndex = 7;
            lbl_VerticalCellAlignments.Text = "Vertical";
            lbl_VerticalCellAlignments.TextAlign = ContentAlignment.BottomCenter;
            // 
            // chk_RowHeight
            // 
            chk_RowHeight.AutoSize = true;
            chk_RowHeight.BackColor = Color.FromArgb(20, 20, 20);
            chk_RowHeight.Checked = true;
            chk_RowHeight.CheckState = CheckState.Checked;
            chk_RowHeight.ForeColor = Color.White;
            chk_RowHeight.Location = new Point(3, 193);
            chk_RowHeight.Name = "chk_RowHeight";
            chk_RowHeight.Size = new Size(96, 18);
            chk_RowHeight.TabIndex = 11;
            chk_RowHeight.Text = "Row Height";
            chk_RowHeight.UseVisualStyleBackColor = false;
            chk_RowHeight.CheckedChanged += chk_RowHeight_CheckedChanged;
            // 
            // tlp_RowHeight
            // 
            tlp_RowHeight.BackColor = Color.FromArgb(40, 40, 40);
            tlp_RowHeight.ColumnCount = 2;
            tlp_RowHeight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_RowHeight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_RowHeight.Controls.Add(label8, 1, 1);
            tlp_RowHeight.Controls.Add(chk_RowHeightAuto, 1, 0);
            tlp_RowHeight.Controls.Add(nud_RowHeight, 0, 0);
            tlp_RowHeight.Controls.Add(nud_RowMaxHeight, 0, 1);
            tlp_RowHeight.Location = new Point(3, 217);
            tlp_RowHeight.Name = "tlp_RowHeight";
            tlp_RowHeight.RowCount = 2;
            tlp_RowHeight.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_RowHeight.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_RowHeight.Size = new Size(645, 56);
            tlp_RowHeight.TabIndex = 16;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.BackColor = Color.FromArgb(40, 40, 40);
            label8.Dock = DockStyle.Fill;
            label8.ForeColor = Color.White;
            label8.Location = new Point(325, 31);
            label8.Margin = new Padding(3);
            label8.Name = "label8";
            label8.Size = new Size(317, 22);
            label8.TabIndex = 2;
            label8.Text = "Max Height";
            label8.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // chk_RowHeightAuto
            // 
            chk_RowHeightAuto.Dock = DockStyle.Fill;
            chk_RowHeightAuto.Location = new Point(325, 3);
            chk_RowHeightAuto.Name = "chk_RowHeightAuto";
            chk_RowHeightAuto.Size = new Size(317, 22);
            chk_RowHeightAuto.TabIndex = 3;
            chk_RowHeightAuto.Text = "Auto";
            chk_RowHeightAuto.UseVisualStyleBackColor = true;
            chk_RowHeightAuto.CheckedChanged += chk_RowHeightAuto_CheckedChanged;
            // 
            // nud_RowHeight
            // 
            nud_RowHeight.BackColor = Color.FromArgb(20, 20, 20);
            nud_RowHeight.DecimalPlaces = 2;
            nud_RowHeight.Dock = DockStyle.Fill;
            nud_RowHeight.ForeColor = Color.White;
            nud_RowHeight.Location = new Point(3, 3);
            nud_RowHeight.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            nud_RowHeight.Name = "nud_RowHeight";
            nud_RowHeight.Size = new Size(316, 22);
            nud_RowHeight.TabIndex = 0;
            nud_RowHeight.TextAlign = HorizontalAlignment.Center;
            nud_RowHeight.Value = new decimal(new int[] { 15, 0, 0, 0 });
            // 
            // nud_RowMaxHeight
            // 
            nud_RowMaxHeight.BackColor = Color.FromArgb(20, 20, 20);
            nud_RowMaxHeight.DecimalPlaces = 2;
            nud_RowMaxHeight.Dock = DockStyle.Fill;
            nud_RowMaxHeight.Enabled = false;
            nud_RowMaxHeight.ForeColor = Color.White;
            nud_RowMaxHeight.Location = new Point(3, 31);
            nud_RowMaxHeight.Maximum = new decimal(new int[] { 30, 0, 0, 0 });
            nud_RowMaxHeight.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
            nud_RowMaxHeight.Name = "nud_RowMaxHeight";
            nud_RowMaxHeight.Size = new Size(316, 22);
            nud_RowMaxHeight.TabIndex = 0;
            nud_RowMaxHeight.TextAlign = HorizontalAlignment.Center;
            nud_RowMaxHeight.Value = new decimal(new int[] { 30, 0, 0, 0 });
            // 
            // chk_ColumnWidth
            // 
            chk_ColumnWidth.AutoSize = true;
            chk_ColumnWidth.BackColor = Color.FromArgb(20, 20, 20);
            chk_ColumnWidth.Checked = true;
            chk_ColumnWidth.CheckState = CheckState.Checked;
            chk_ColumnWidth.ForeColor = Color.White;
            chk_ColumnWidth.Location = new Point(3, 279);
            chk_ColumnWidth.Name = "chk_ColumnWidth";
            chk_ColumnWidth.Size = new Size(110, 18);
            chk_ColumnWidth.TabIndex = 13;
            chk_ColumnWidth.Text = "Column Width";
            chk_ColumnWidth.UseVisualStyleBackColor = false;
            chk_ColumnWidth.CheckedChanged += chk_ColumnWidth_CheckedChanged;
            // 
            // tlp_ColumnWidth
            // 
            tlp_ColumnWidth.BackColor = Color.FromArgb(40, 40, 40);
            tlp_ColumnWidth.ColumnCount = 2;
            tlp_ColumnWidth.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_ColumnWidth.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_ColumnWidth.Controls.Add(label9, 1, 1);
            tlp_ColumnWidth.Controls.Add(chk_ColumnWidthAuto, 1, 0);
            tlp_ColumnWidth.Controls.Add(nud_ColumnWidth, 0, 0);
            tlp_ColumnWidth.Controls.Add(nud_ColumnMaxWidth, 0, 1);
            tlp_ColumnWidth.Location = new Point(3, 303);
            tlp_ColumnWidth.Name = "tlp_ColumnWidth";
            tlp_ColumnWidth.RowCount = 2;
            tlp_ColumnWidth.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_ColumnWidth.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_ColumnWidth.Size = new Size(645, 56);
            tlp_ColumnWidth.TabIndex = 16;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.BackColor = Color.FromArgb(40, 40, 40);
            label9.Dock = DockStyle.Fill;
            label9.ForeColor = Color.White;
            label9.Location = new Point(325, 31);
            label9.Margin = new Padding(3);
            label9.Name = "label9";
            label9.Size = new Size(317, 22);
            label9.TabIndex = 3;
            label9.Text = "Max Width";
            label9.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // chk_ColumnWidthAuto
            // 
            chk_ColumnWidthAuto.AutoSize = true;
            chk_ColumnWidthAuto.Dock = DockStyle.Fill;
            chk_ColumnWidthAuto.Location = new Point(325, 3);
            chk_ColumnWidthAuto.Name = "chk_ColumnWidthAuto";
            chk_ColumnWidthAuto.Size = new Size(317, 22);
            chk_ColumnWidthAuto.TabIndex = 4;
            chk_ColumnWidthAuto.Text = "Auto";
            chk_ColumnWidthAuto.UseVisualStyleBackColor = true;
            chk_ColumnWidthAuto.CheckedChanged += chk_ColumnWidthAuto_CheckedChanged;
            // 
            // nud_ColumnWidth
            // 
            nud_ColumnWidth.BackColor = Color.FromArgb(20, 20, 20);
            nud_ColumnWidth.DecimalPlaces = 2;
            nud_ColumnWidth.Dock = DockStyle.Fill;
            nud_ColumnWidth.ForeColor = Color.White;
            nud_ColumnWidth.Location = new Point(3, 3);
            nud_ColumnWidth.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            nud_ColumnWidth.Name = "nud_ColumnWidth";
            nud_ColumnWidth.Size = new Size(316, 22);
            nud_ColumnWidth.TabIndex = 0;
            nud_ColumnWidth.TextAlign = HorizontalAlignment.Center;
            nud_ColumnWidth.Value = new decimal(new int[] { 20, 0, 0, 0 });
            // 
            // nud_ColumnMaxWidth
            // 
            nud_ColumnMaxWidth.BackColor = Color.FromArgb(20, 20, 20);
            nud_ColumnMaxWidth.DecimalPlaces = 2;
            nud_ColumnMaxWidth.Dock = DockStyle.Fill;
            nud_ColumnMaxWidth.Enabled = false;
            nud_ColumnMaxWidth.ForeColor = Color.White;
            nud_ColumnMaxWidth.Location = new Point(3, 31);
            nud_ColumnMaxWidth.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
            nud_ColumnMaxWidth.Name = "nud_ColumnMaxWidth";
            nud_ColumnMaxWidth.Size = new Size(316, 22);
            nud_ColumnMaxWidth.TabIndex = 0;
            nud_ColumnMaxWidth.TextAlign = HorizontalAlignment.Center;
            nud_ColumnMaxWidth.Value = new decimal(new int[] { 50, 0, 0, 0 });
            // 
            // chk_Zoom
            // 
            chk_Zoom.AutoSize = true;
            chk_Zoom.BackColor = Color.FromArgb(20, 20, 20);
            chk_Zoom.Checked = true;
            chk_Zoom.CheckState = CheckState.Checked;
            chk_Zoom.ForeColor = Color.White;
            chk_Zoom.Location = new Point(3, 365);
            chk_Zoom.Name = "chk_Zoom";
            chk_Zoom.Size = new Size(54, 18);
            chk_Zoom.TabIndex = 15;
            chk_Zoom.Text = "Zoom";
            chk_Zoom.UseVisualStyleBackColor = false;
            chk_Zoom.CheckedChanged += chk_Zoom_CheckedChanged;
            // 
            // tlp_Zoom
            // 
            tlp_Zoom.BackColor = Color.FromArgb(40, 40, 40);
            tlp_Zoom.ColumnCount = 2;
            tlp_Zoom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_Zoom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp_Zoom.Controls.Add(nud_Zoom, 0, 0);
            tlp_Zoom.Controls.Add(label10, 1, 0);
            tlp_Zoom.Location = new Point(3, 389);
            tlp_Zoom.Name = "tlp_Zoom";
            tlp_Zoom.RowCount = 1;
            tlp_Zoom.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_Zoom.Size = new Size(645, 28);
            tlp_Zoom.TabIndex = 16;
            // 
            // nud_Zoom
            // 
            nud_Zoom.BackColor = Color.FromArgb(20, 20, 20);
            nud_Zoom.Dock = DockStyle.Fill;
            nud_Zoom.ForeColor = Color.White;
            nud_Zoom.Increment = new decimal(new int[] { 10, 0, 0, 0 });
            nud_Zoom.Location = new Point(3, 3);
            nud_Zoom.Maximum = new decimal(new int[] { 400, 0, 0, 0 });
            nud_Zoom.Minimum = new decimal(new int[] { 10, 0, 0, 0 });
            nud_Zoom.Name = "nud_Zoom";
            nud_Zoom.Size = new Size(316, 22);
            nud_Zoom.TabIndex = 0;
            nud_Zoom.TextAlign = HorizontalAlignment.Center;
            nud_Zoom.Value = new decimal(new int[] { 100, 0, 0, 0 });
            // 
            // label10
            // 
            label10.BackColor = Color.FromArgb(40, 40, 40);
            label10.Dock = DockStyle.Fill;
            label10.ForeColor = Color.White;
            label10.Location = new Point(325, 0);
            label10.Name = "label10";
            label10.Size = new Size(317, 28);
            label10.TabIndex = 2;
            label10.Text = "%";
            label10.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.BackColor = Color.FromArgb(20, 20, 20);
            tableLayoutPanel3.ColumnCount = 2;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.Controls.Add(btn_Start, 0, 0);
            tableLayoutPanel3.Dock = DockStyle.Fill;
            tableLayoutPanel3.ForeColor = Color.White;
            tableLayoutPanel3.Location = new Point(0, 499);
            tableLayoutPanel3.Margin = new Padding(0);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.Size = new Size(684, 50);
            tableLayoutPanel3.TabIndex = 11;
            // 
            // tlp_Main
            // 
            tlp_Main.BackColor = Color.FromArgb(20, 20, 20);
            tlp_Main.ColumnCount = 1;
            tlp_Main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlp_Main.Controls.Add(tableLayoutPanel1, 0, 0);
            tlp_Main.Controls.Add(tableLayoutPanel3, 0, 2);
            tlp_Main.Controls.Add(tableLayoutPanel2, 0, 1);
            tlp_Main.Controls.Add(lbl_Credits, 0, 3);
            tlp_Main.Dock = DockStyle.Fill;
            tlp_Main.ForeColor = Color.White;
            tlp_Main.Location = new Point(0, 0);
            tlp_Main.Margin = new Padding(0);
            tlp_Main.Name = "tlp_Main";
            tlp_Main.RowCount = 4;
            tlp_Main.RowStyles.Add(new RowStyle(SizeType.Absolute, 106F));
            tlp_Main.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_Main.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            tlp_Main.RowStyles.Add(new RowStyle(SizeType.Absolute, 12F));
            tlp_Main.Size = new Size(684, 561);
            tlp_Main.TabIndex = 12;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 3;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33333F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3333359F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3333359F));
            tableLayoutPanel2.Controls.Add(tableLayoutPanel5, 0, 0);
            tableLayoutPanel2.Controls.Add(pnl_Options, 0, 2);
            tableLayoutPanel2.Controls.Add(tableLayoutPanel4, 0, 1);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 109);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 3;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Absolute, 43F));
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(678, 387);
            tableLayoutPanel2.TabIndex = 13;
            // 
            // tableLayoutPanel5
            // 
            tableLayoutPanel5.ColumnCount = 12;
            tableLayoutPanel2.SetColumnSpan(tableLayoutPanel5, 3);
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333335F));
            tableLayoutPanel5.Controls.Add(cmb_Preset, 0, 1);
            tableLayoutPanel5.Controls.Add(lbl_Preset, 0, 0);
            tableLayoutPanel5.Controls.Add(btn_SavePreset, 10, 1);
            tableLayoutPanel5.Controls.Add(btn_RemovePreset, 11, 1);
            tableLayoutPanel5.Dock = DockStyle.Fill;
            tableLayoutPanel5.Location = new Point(0, 0);
            tableLayoutPanel5.Margin = new Padding(0);
            tableLayoutPanel5.Name = "tableLayoutPanel5";
            tableLayoutPanel5.RowCount = 2;
            tableLayoutPanel5.RowStyles.Add(new RowStyle(SizeType.Absolute, 18F));
            tableLayoutPanel5.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel5.Size = new Size(678, 43);
            tableLayoutPanel5.TabIndex = 5;
            // 
            // cmb_Preset
            // 
            cmb_Preset.BackColor = Color.FromArgb(40, 40, 40);
            tableLayoutPanel5.SetColumnSpan(cmb_Preset, 10);
            cmb_Preset.Dock = DockStyle.Fill;
            cmb_Preset.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_Preset.FlatStyle = FlatStyle.Flat;
            cmb_Preset.ForeColor = SystemColors.HighlightText;
            cmb_Preset.FormattingEnabled = true;
            cmb_Preset.Items.AddRange(new object[] { "Default" });
            cmb_Preset.Location = new Point(3, 18);
            cmb_Preset.Margin = new Padding(3, 0, 3, 3);
            cmb_Preset.Name = "cmb_Preset";
            cmb_Preset.Size = new Size(554, 22);
            cmb_Preset.TabIndex = 0;
            cmb_Preset.SelectedIndexChanged += cmb_Preset_SelectedIndexChanged;
            // 
            // lbl_Preset
            // 
            lbl_Preset.AutoSize = true;
            tableLayoutPanel5.SetColumnSpan(lbl_Preset, 6);
            lbl_Preset.Dock = DockStyle.Fill;
            lbl_Preset.Font = new Font("Consolas", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            lbl_Preset.Location = new Point(3, 0);
            lbl_Preset.Name = "lbl_Preset";
            lbl_Preset.Size = new Size(330, 18);
            lbl_Preset.TabIndex = 1;
            lbl_Preset.Text = "Preset";
            // 
            // btn_SavePreset
            // 
            btn_SavePreset.FlatStyle = FlatStyle.Flat;
            btn_SavePreset.Location = new Point(563, 18);
            btn_SavePreset.Margin = new Padding(3, 0, 3, 0);
            btn_SavePreset.Name = "btn_SavePreset";
            btn_SavePreset.Size = new Size(50, 23);
            btn_SavePreset.TabIndex = 2;
            btn_SavePreset.Text = "S";
            btn_SavePreset.UseVisualStyleBackColor = true;
            // 
            // btn_RemovePreset
            // 
            btn_RemovePreset.FlatStyle = FlatStyle.Flat;
            btn_RemovePreset.Location = new Point(619, 18);
            btn_RemovePreset.Margin = new Padding(3, 0, 3, 0);
            btn_RemovePreset.Name = "btn_RemovePreset";
            btn_RemovePreset.Size = new Size(56, 23);
            btn_RemovePreset.TabIndex = 3;
            btn_RemovePreset.Text = "R";
            btn_RemovePreset.UseVisualStyleBackColor = true;
            // 
            // pnl_Options
            // 
            tableLayoutPanel2.SetColumnSpan(pnl_Options, 3);
            pnl_Options.Controls.Add(pnl_Apply);
            pnl_Options.Controls.Add(pnl_Remove);
            pnl_Options.Controls.Add(pnl_Others);
            pnl_Options.Dock = DockStyle.Fill;
            pnl_Options.Location = new Point(3, 76);
            pnl_Options.Name = "pnl_Options";
            pnl_Options.Size = new Size(672, 308);
            pnl_Options.TabIndex = 0;
            // 
            // pnl_Apply
            // 
            pnl_Apply.Controls.Add(flp_Apply);
            pnl_Apply.Dock = DockStyle.Fill;
            pnl_Apply.Location = new Point(0, 0);
            pnl_Apply.Name = "pnl_Apply";
            pnl_Apply.Size = new Size(672, 308);
            pnl_Apply.TabIndex = 0;
            // 
            // pnl_Remove
            // 
            pnl_Remove.Controls.Add(flp_Remove);
            pnl_Remove.Dock = DockStyle.Fill;
            pnl_Remove.Location = new Point(0, 0);
            pnl_Remove.Name = "pnl_Remove";
            pnl_Remove.Size = new Size(672, 308);
            pnl_Remove.TabIndex = 1;
            // 
            // pnl_Others
            // 
            pnl_Others.Controls.Add(flp_Others);
            pnl_Others.Dock = DockStyle.Fill;
            pnl_Others.Location = new Point(0, 0);
            pnl_Others.Name = "pnl_Others";
            pnl_Others.Size = new Size(672, 308);
            pnl_Others.TabIndex = 2;
            // 
            // flp_Others
            // 
            flp_Others.AutoScroll = true;
            flp_Others.AutoSize = true;
            flp_Others.BackColor = Color.FromArgb(20, 20, 20);
            flp_Others.BorderStyle = BorderStyle.FixedSingle;
            flp_Others.Controls.Add(chk_GetLastRealEmptyRow);
            flp_Others.Controls.Add(chk_GetLastRealEmptyColumn);
            flp_Others.Controls.Add(chk_FindHeader);
            flp_Others.Controls.Add(tlp_FindHeader);
            flp_Others.Dock = DockStyle.Fill;
            flp_Others.FlowDirection = FlowDirection.TopDown;
            flp_Others.ForeColor = Color.White;
            flp_Others.Location = new Point(0, 0);
            flp_Others.Name = "flp_Others";
            flp_Others.Size = new Size(672, 308);
            flp_Others.TabIndex = 11;
            flp_Others.WrapContents = false;
            // 
            // chk_GetLastRealEmptyRow
            // 
            chk_GetLastRealEmptyRow.AutoSize = true;
            chk_GetLastRealEmptyRow.Checked = true;
            chk_GetLastRealEmptyRow.CheckState = CheckState.Checked;
            chk_GetLastRealEmptyRow.Location = new Point(3, 3);
            chk_GetLastRealEmptyRow.Name = "chk_GetLastRealEmptyRow";
            chk_GetLastRealEmptyRow.Size = new Size(187, 18);
            chk_GetLastRealEmptyRow.TabIndex = 1;
            chk_GetLastRealEmptyRow.Text = "Get Last Real Empty Row";
            chk_GetLastRealEmptyRow.UseVisualStyleBackColor = true;
            // 
            // chk_GetLastRealEmptyColumn
            // 
            chk_GetLastRealEmptyColumn.AutoSize = true;
            chk_GetLastRealEmptyColumn.Checked = true;
            chk_GetLastRealEmptyColumn.CheckState = CheckState.Checked;
            chk_GetLastRealEmptyColumn.Location = new Point(3, 27);
            chk_GetLastRealEmptyColumn.Name = "chk_GetLastRealEmptyColumn";
            chk_GetLastRealEmptyColumn.Size = new Size(208, 18);
            chk_GetLastRealEmptyColumn.TabIndex = 0;
            chk_GetLastRealEmptyColumn.Text = "Get Last Real Empty Column";
            chk_GetLastRealEmptyColumn.UseVisualStyleBackColor = true;
            // 
            // chk_FindHeader
            // 
            chk_FindHeader.AutoSize = true;
            chk_FindHeader.Checked = true;
            chk_FindHeader.CheckState = CheckState.Checked;
            chk_FindHeader.Location = new Point(3, 51);
            chk_FindHeader.Name = "chk_FindHeader";
            chk_FindHeader.Size = new Size(103, 18);
            chk_FindHeader.TabIndex = 18;
            chk_FindHeader.Text = "Find Header";
            chk_FindHeader.UseVisualStyleBackColor = true;
            chk_FindHeader.CheckedChanged += chk_FindHeader_CheckedChanged;
            // 
            // tlp_FindHeader
            // 
            tlp_FindHeader.BackColor = Color.FromArgb(40, 40, 40);
            tlp_FindHeader.ColumnCount = 10;
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
            tlp_FindHeader.Controls.Add(cmb_FindHeaderFilterOption, 0, 0);
            tlp_FindHeader.Controls.Add(txt_FindHeader, 2, 0);
            tlp_FindHeader.Controls.Add(btn_FindHeaderAdd, 8, 0);
            tlp_FindHeader.Controls.Add(btn_FindHeaderRemove, 9, 0);
            tlp_FindHeader.Controls.Add(pnl_FindHeader, 0, 1);
            tlp_FindHeader.Location = new Point(3, 75);
            tlp_FindHeader.Name = "tlp_FindHeader";
            tlp_FindHeader.RowCount = 5;
            tlp_FindHeader.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp_FindHeader.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_FindHeader.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_FindHeader.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_FindHeader.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            tlp_FindHeader.Size = new Size(645, 140);
            tlp_FindHeader.TabIndex = 17;
            // 
            // cmb_FindHeaderFilterOption
            // 
            cmb_FindHeaderFilterOption.BackColor = Color.FromArgb(20, 20, 20);
            tlp_FindHeader.SetColumnSpan(cmb_FindHeaderFilterOption, 2);
            cmb_FindHeaderFilterOption.Dock = DockStyle.Fill;
            cmb_FindHeaderFilterOption.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb_FindHeaderFilterOption.FlatStyle = FlatStyle.Flat;
            cmb_FindHeaderFilterOption.ForeColor = Color.FromArgb(224, 224, 224);
            cmb_FindHeaderFilterOption.FormattingEnabled = true;
            cmb_FindHeaderFilterOption.Items.AddRange(new object[] { "Equals", "Contains", "Starts With", "Ends With" });
            cmb_FindHeaderFilterOption.Location = new Point(3, 3);
            cmb_FindHeaderFilterOption.Name = "cmb_FindHeaderFilterOption";
            cmb_FindHeaderFilterOption.Size = new Size(122, 22);
            cmb_FindHeaderFilterOption.TabIndex = 5;
            // 
            // txt_FindHeader
            // 
            txt_FindHeader.BackColor = Color.FromArgb(20, 20, 20);
            txt_FindHeader.BorderStyle = BorderStyle.FixedSingle;
            tlp_FindHeader.SetColumnSpan(txt_FindHeader, 6);
            txt_FindHeader.Dock = DockStyle.Fill;
            txt_FindHeader.ForeColor = Color.FromArgb(224, 224, 224);
            txt_FindHeader.Location = new Point(131, 3);
            txt_FindHeader.Name = "txt_FindHeader";
            txt_FindHeader.PlaceholderText = "Enter your text here";
            txt_FindHeader.Size = new Size(378, 22);
            txt_FindHeader.TabIndex = 3;
            // 
            // btn_FindHeaderAdd
            // 
            btn_FindHeaderAdd.BackColor = Color.FromArgb(40, 40, 40);
            btn_FindHeaderAdd.Dock = DockStyle.Fill;
            btn_FindHeaderAdd.FlatStyle = FlatStyle.Flat;
            btn_FindHeaderAdd.Location = new Point(515, 3);
            btn_FindHeaderAdd.Name = "btn_FindHeaderAdd";
            btn_FindHeaderAdd.Size = new Size(58, 22);
            btn_FindHeaderAdd.TabIndex = 0;
            btn_FindHeaderAdd.Text = "+";
            btn_FindHeaderAdd.UseVisualStyleBackColor = false;
            btn_FindHeaderAdd.Click += btn_FindHeaderAdd_Click;
            // 
            // btn_FindHeaderRemove
            // 
            btn_FindHeaderRemove.BackColor = Color.FromArgb(40, 40, 40);
            btn_FindHeaderRemove.Dock = DockStyle.Fill;
            btn_FindHeaderRemove.FlatStyle = FlatStyle.Flat;
            btn_FindHeaderRemove.Location = new Point(579, 3);
            btn_FindHeaderRemove.Name = "btn_FindHeaderRemove";
            btn_FindHeaderRemove.Size = new Size(63, 22);
            btn_FindHeaderRemove.TabIndex = 1;
            btn_FindHeaderRemove.Text = "-";
            btn_FindHeaderRemove.UseVisualStyleBackColor = false;
            btn_FindHeaderRemove.Click += btn_FindHeaderRemove_Click;
            // 
            // pnl_FindHeader
            // 
            pnl_FindHeader.BackColor = Color.FromArgb(20, 20, 20);
            pnl_FindHeader.BorderStyle = BorderStyle.FixedSingle;
            tlp_FindHeader.SetColumnSpan(pnl_FindHeader, 10);
            pnl_FindHeader.Controls.Add(lst_FindHeader);
            pnl_FindHeader.Dock = DockStyle.Fill;
            pnl_FindHeader.Location = new Point(3, 31);
            pnl_FindHeader.Name = "pnl_FindHeader";
            tlp_FindHeader.SetRowSpan(pnl_FindHeader, 4);
            pnl_FindHeader.Size = new Size(639, 106);
            pnl_FindHeader.TabIndex = 4;
            // 
            // lst_FindHeader
            // 
            lst_FindHeader.BackColor = Color.FromArgb(20, 20, 20);
            lst_FindHeader.BorderStyle = BorderStyle.None;
            lst_FindHeader.Dock = DockStyle.Fill;
            lst_FindHeader.Font = new Font("Consolas", 8.25F, FontStyle.Regular, GraphicsUnit.Point);
            lst_FindHeader.ForeColor = Color.FromArgb(224, 224, 224);
            lst_FindHeader.FormattingEnabled = true;
            lst_FindHeader.Items.AddRange(new object[] { "chapa", "cnpj", "colaborador", "comprafinal", "cpf", "cunid", "datade", "datanasc", "depto", "desc", "diasut", "emissor", "empresa", "escala", "funcao", "funcionari", "mat", "nascimento", "ncartao", "nome", "nrcartao", "nrdocartao", "numero", "operadora", "parcela", "passage", "qtded", "quantidade", "quinzena", "qvt", "razaosocial", "rg", "saldo", "sexo", "total", "tvt", "uf", "valordia", "valordias", "vtdia", "vttotal", "vvt" });
            lst_FindHeader.Location = new Point(0, 0);
            lst_FindHeader.Name = "lst_FindHeader";
            lst_FindHeader.Size = new Size(637, 104);
            lst_FindHeader.Sorted = true;
            lst_FindHeader.TabIndex = 2;
            // 
            // tableLayoutPanel4
            // 
            tableLayoutPanel4.ColumnCount = 12;
            tableLayoutPanel2.SetColumnSpan(tableLayoutPanel4, 3);
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8.333333F));
            tableLayoutPanel4.Controls.Add(btn_Apply, 0, 0);
            tableLayoutPanel4.Controls.Add(btn_Remove, 4, 0);
            tableLayoutPanel4.Controls.Add(btn_Others, 8, 0);
            tableLayoutPanel4.Dock = DockStyle.Fill;
            tableLayoutPanel4.Location = new Point(0, 43);
            tableLayoutPanel4.Margin = new Padding(0);
            tableLayoutPanel4.Name = "tableLayoutPanel4";
            tableLayoutPanel4.RowCount = 1;
            tableLayoutPanel4.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel4.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel4.Size = new Size(678, 30);
            tableLayoutPanel4.TabIndex = 4;
            // 
            // btn_Apply
            // 
            btn_Apply.BackColor = Color.FromArgb(20, 20, 20);
            tableLayoutPanel4.SetColumnSpan(btn_Apply, 4);
            btn_Apply.Dock = DockStyle.Fill;
            btn_Apply.FlatStyle = FlatStyle.Flat;
            btn_Apply.Location = new Point(3, 3);
            btn_Apply.Name = "btn_Apply";
            btn_Apply.Size = new Size(218, 24);
            btn_Apply.TabIndex = 1;
            btn_Apply.Text = "Apply";
            btn_Apply.UseVisualStyleBackColor = false;
            btn_Apply.Click += btn_Apply_Click;
            // 
            // btn_Remove
            // 
            btn_Remove.BackColor = Color.FromArgb(20, 20, 20);
            tableLayoutPanel4.SetColumnSpan(btn_Remove, 4);
            btn_Remove.Dock = DockStyle.Fill;
            btn_Remove.FlatStyle = FlatStyle.Flat;
            btn_Remove.Location = new Point(227, 3);
            btn_Remove.Name = "btn_Remove";
            btn_Remove.Size = new Size(218, 24);
            btn_Remove.TabIndex = 2;
            btn_Remove.Text = "Remove";
            btn_Remove.UseVisualStyleBackColor = false;
            btn_Remove.Click += btn_Remove_Click;
            // 
            // btn_Others
            // 
            btn_Others.BackColor = Color.FromArgb(20, 20, 20);
            tableLayoutPanel4.SetColumnSpan(btn_Others, 4);
            btn_Others.Dock = DockStyle.Fill;
            btn_Others.FlatStyle = FlatStyle.Flat;
            btn_Others.Location = new Point(451, 3);
            btn_Others.Name = "btn_Others";
            btn_Others.Size = new Size(224, 24);
            btn_Others.TabIndex = 3;
            btn_Others.Text = "Others";
            btn_Others.UseVisualStyleBackColor = false;
            btn_Others.Click += btn_Others_Click;
            // 
            // lbl_Credits
            // 
            lbl_Credits.AutoSize = true;
            lbl_Credits.Cursor = Cursors.Hand;
            lbl_Credits.Dock = DockStyle.Fill;
            lbl_Credits.Font = new Font("Consolas", 8.25F, FontStyle.Bold, GraphicsUnit.Point);
            lbl_Credits.Location = new Point(3, 549);
            lbl_Credits.Name = "lbl_Credits";
            lbl_Credits.Size = new Size(678, 12);
            lbl_Credits.TabIndex = 14;
            lbl_Credits.Text = "Developed by GCScript";
            lbl_Credits.TextAlign = ContentAlignment.MiddleCenter;
            lbl_Credits.Click += lbl_Credits_Click;
            // 
            // frm_Main
            // 
            AutoScaleDimensions = new SizeF(7F, 14F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(24, 26, 27);
            ClientSize = new Size(684, 561);
            Controls.Add(tlp_Main);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MaximumSize = new Size(700, 600);
            MinimumSize = new Size(700, 600);
            Name = "frm_Main";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "GCScript Excel Tools";
            Load += frm_Main_Load;
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.PerformLayout();
            panel1.ResumeLayout(false);
            flp_Remove.ResumeLayout(false);
            flp_Remove.PerformLayout();
            tlp_RemoveColumns.ResumeLayout(false);
            tlp_RemoveColumns.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgv_RemoveColumns).EndInit();
            flp_Apply.ResumeLayout(false);
            flp_Apply.PerformLayout();
            tlp_SortWorksheets.ResumeLayout(false);
            tlp_Font.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)nud_FontSize).EndInit();
            tlp_CellAlignments.ResumeLayout(false);
            tlp_RowHeight.ResumeLayout(false);
            tlp_RowHeight.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)nud_RowHeight).EndInit();
            ((System.ComponentModel.ISupportInitialize)nud_RowMaxHeight).EndInit();
            tlp_ColumnWidth.ResumeLayout(false);
            tlp_ColumnWidth.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)nud_ColumnWidth).EndInit();
            ((System.ComponentModel.ISupportInitialize)nud_ColumnMaxWidth).EndInit();
            tlp_Zoom.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)nud_Zoom).EndInit();
            tableLayoutPanel3.ResumeLayout(false);
            tlp_Main.ResumeLayout(false);
            tlp_Main.PerformLayout();
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel5.ResumeLayout(false);
            tableLayoutPanel5.PerformLayout();
            pnl_Options.ResumeLayout(false);
            pnl_Apply.ResumeLayout(false);
            pnl_Apply.PerformLayout();
            pnl_Remove.ResumeLayout(false);
            pnl_Remove.PerformLayout();
            pnl_Others.ResumeLayout(false);
            pnl_Others.PerformLayout();
            flp_Others.ResumeLayout(false);
            flp_Others.PerformLayout();
            tlp_FindHeader.ResumeLayout(false);
            tlp_FindHeader.PerformLayout();
            pnl_FindHeader.ResumeLayout(false);
            tableLayoutPanel4.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion
        private Button btn_Add;
        private ListBox lst_FilesPath;
        private TableLayoutPanel tableLayoutPanel1;
        private Button btn_RemoveFile;
        private Label label1;
        private Panel panel1;
        private Button btn_Start;
        private Label lbl_Others;
        private FlowLayoutPanel flp_Apply;
        private CheckBox chk_SortWorksheets;
        private CheckBox chk_RemoveEmptyWorksheets;
        private CheckBox chk_RemoveBackgroundColor;
        private CheckBox chk_RemoveFontColor;
        private CheckBox chk_RemoveFormatting;
        private CheckBox chk_RemoveEmptyColumns;
        private CheckBox chk_RemoveEmptyRows;
        private NumericUpDown nud_FontSize;
        private ComboBox cmb_FontName;
        private CheckBox chk_Font;
        private FlowLayoutPanel flp_Remove;
        private CheckBox chk_ColumnWidthAuto;
        private CheckBox chk_CellAlignments;
        private CheckBox chk_RowHeight;
        private CheckBox chk_ColumnWidth;
        private CheckBox chk_Zoom;
        private Label lbl_HorizontalCellAlignments;
        private ComboBox cmb_VerticalCellAlignments;
        private Label lbl_VerticalCellAlignments;
        private ComboBox cmb_HorizontalCellAlignments;
        private NumericUpDown nud_RowHeight;
        private NumericUpDown nud_RowMaxHeight;
        private NumericUpDown nud_ColumnWidth;
        private NumericUpDown nud_ColumnMaxWidth;
        private Label label8;
        private Label label9;
        private ComboBox cmb_SortWorksheets;
        private Label label10;
        private NumericUpDown nud_Zoom;
        private TableLayoutPanel tableLayoutPanel3;
        private TableLayoutPanel tlp_Main;
        private CheckBox chk_RemoveInvisibleWorksheets;
        private TableLayoutPanel tableLayoutPanel2;
        private Panel pnl_Options;
        private Panel pnl_Remove;
        private Panel pnl_Apply;
        private Button btn_Apply;
        private Button btn_Remove;
        private Button btn_Others;
        private CheckBox chk_GetLastRealEmptyColumn;
        private CheckBox chk_RowHeightAuto;
        private TableLayoutPanel tlp_RowHeight;
        private TableLayoutPanel tlp_ColumnWidth;
        private TableLayoutPanel tlp_CellAlignments;
        private TableLayoutPanel tlp_Font;
        private TableLayoutPanel tlp_SortWorksheets;
        private TableLayoutPanel tlp_Zoom;
        private CheckBox chk_GetLastRealEmptyRow;
        private Panel pnl_Others;
        private FlowLayoutPanel flp_Others;
        private Label lbl_Credits;
        private CheckBox chk_FindHeader;
        private TableLayoutPanel tlp_FindHeader;
        private Button btn_FindHeaderAdd;
        private Button btn_FindHeaderRemove;
        private ListBox lst_FindHeader;
        private TextBox txt_FindHeader;
        private Panel pnl_FindHeader;
        private CheckBox chk_RemoveColumns;
        private TableLayoutPanel tlp_RemoveColumns;
        private Button btn_RemoveColumnsAdd;
        private Button btn_RemoveColumnsRemove;
        private TextBox txt_RemoveColumns;
        private ComboBox cmb_FindHeaderFilterOption;
        private ComboBox cmb_RemoveColumnsFilterOption;
        private DataGridView dgv_RemoveColumns;
        private CheckBox chk_RemoveHiddenRows;
        private TableLayoutPanel tableLayoutPanel5;
        private ComboBox cmb_Preset;
        private Label lbl_Preset;
        private TableLayoutPanel tableLayoutPanel4;
        private Button btn_SavePreset;
        private Button btn_RemovePreset;
    }
}