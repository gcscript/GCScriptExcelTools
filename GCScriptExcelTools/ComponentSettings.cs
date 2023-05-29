namespace GCScriptExcelTools;

public static class ComponentSettings
{
    public enum EDataGridViewCellTemplate { TextBox, ComboBox, CheckBox, }

    private static void DataGridViewDarkMode(DataGridView dataGridView)
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

        dataGridView.DefaultCellStyle.Font = new Font("Consolas", 8);
        dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font("Consolas", 9, FontStyle.Bold); // Define o estilo da fonte do cabeçalho
        dataGridView.RowHeadersVisible = false; // Define a coluna do cabeçalho como invisível
        dataGridView.BackgroundColor = BackColor; // Define a cor de fundo do DataGridView
        dataGridView.DefaultCellStyle.BackColor = BackColor; // Define a cor de fundo das células para uma cor escura
        dataGridView.DefaultCellStyle.ForeColor = ForeColor; // Define a cor do texto das células para branco
        dataGridView.DefaultCellStyle.SelectionBackColor = SelectionBackColor; // Define a cor de fundo das células selecionadas para uma cor escura
        dataGridView.DefaultCellStyle.SelectionForeColor = SelectionForeColor; // Define a cor do texto das células selecionadas para branco
        dataGridView.ColumnHeadersDefaultCellStyle.BackColor = ColumnHeadersBackColor; // Define a cor de fundo do cabeçalho para uma cor escura
        dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = ColumnHeadersForeColor; // Define a cor do texto do cabeçalho para branco
        dataGridView.ColumnHeadersDefaultCellStyle.SelectionBackColor = ColumnHeadersSelectionBackColor; // Define a cor de fundo da célula do cabeçalho quando selecionada para uma cor escura
        dataGridView.ColumnHeadersDefaultCellStyle.SelectionForeColor = ColumnHeadersSelectionForeColor; // Define a cor do texto da célula do cabeçalho quando selecionada para branco
        dataGridView.GridColor = GridColor; // Define a cor da borda do DataGridView
        dataGridView.EnableHeadersVisualStyles = false; // Define se o estilo padrão do cabeçalho será usado
        dataGridView.BorderStyle = BorderStyle.None; // Define se a borda do DataGridView será exibida
        dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.None; // Define o estilo da borda das células
        dataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; // Define o estilo da borda do cabeçalho
        dataGridView.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; // Define o estilo da borda do cabeçalho da linha
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect; // Define o modo de seleção
        dataGridView.MultiSelect = false; // Define se o usuário pode selecionar mais de uma linha
        dataGridView.AllowUserToOrderColumns = true; // Define se o usuário pode ordenar as colunas;
        dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing; // Define se o usuário pode redimensionar a altura do cabeçalho

        dataGridView.AllowUserToAddRows = false;
        dataGridView.AllowUserToDeleteRows = false;
        dataGridView.AllowUserToResizeRows = false;
        dataGridView.ReadOnly = true;
    }
    public static void CreateDataGridView(DataGridView dataGridView, Panel destiny, bool darkMode = true)
    {
        dataGridView.Dock = DockStyle.Fill;
        if (darkMode)
        {
            DataGridViewDarkMode(dataGridView);
        }
        destiny.Controls.Add(dataGridView);
    }


    public static void CreateColumnDataGridView(DataGridView dataGridView, string name, string header, EDataGridViewCellTemplate cellTemplate, DataGridViewAutoSizeColumnMode autoSizeMode)
    {
        DataGridViewColumn column = new()
        {
            Name = name,
            HeaderText = header,
            ReadOnly = false,
            Visible = true,
            SortMode = DataGridViewColumnSortMode.Automatic,

            CellTemplate = cellTemplate switch
            {
                EDataGridViewCellTemplate.TextBox => new DataGridViewTextBoxCell(),
                EDataGridViewCellTemplate.ComboBox => new DataGridViewComboBoxCell(),
                EDataGridViewCellTemplate.CheckBox => new DataGridViewCheckBoxCell(),
                _ => new DataGridViewTextBoxCell(),
            },

            AutoSizeMode = autoSizeMode,
        };

        dataGridView.Columns.Add(column);
    }

    public static bool DataGridViewCheckIfTheValueExists(DataGridView dataGridView, string columnName, string value)
    {
        try
        {
            value = Tools.ProcessTextForComparison(text: value);

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                string currentValue = Tools.ProcessTextForComparison(text: row.Cells[columnName].Value.ToString()!);

                if (currentValue == value)
                {
                    return true;
                }
            }
            return false;
        }
        catch
        {
            return true;
        }
    }
}
