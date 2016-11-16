using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Xml;

namespace IF97Tool
{
    public partial class IF97ToolForm : Form
    {
        public static int editingRow = 0;
        public static int editingColumn = 0;
        public static int selectedRow = 0;
        public static int selectedColumn = 0;
        public static List<int> selectedRows = new List<int>();
        public static bool editing = false;
        public static bool updating = false;
        public IF97ToolForm()
        {
            
            InitializeComponent();
        }

        private void IF97ToolForm_Load(object sender, EventArgs e)
        {

            // TODO: This line of code loads data into the 'if97VersionExpDataSet.IFLoad' table. You can move, or remove it, as needed.
        }

        private void IF97ToolForm_Shown(object sender, EventArgs e)
        {
            try
            {
                this.iFLoadTableAdapter._connection.ConnectionString = Properties.Settings.Default.If97VersionExpConnectionString;
                this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                dataSyncTimer.Start();
            }
            catch
            {
                try
                {
                    var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    var connectionStringsSection = (ConnectionStringsSection)config.GetSection("connectionStrings");
                    var connectionString = connectionStringsSection.ConnectionStrings["IF97Tool.Properties.Settings.If97VersionExpConnectionString"].ConnectionString;
                    Debug.WriteLine("Original connectionString: " + connectionString);
                    string configDataSource = new System.Data.OleDb.OleDbConnectionStringBuilder(connectionString).DataSource;
                    string dataSource = GetDataSource(configDataSource);
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataSource;
                    connectionStringsSection.ConnectionStrings["IF97Tool.Properties.Settings.If97VersionExpConnectionString"].ConnectionString = connectionString;
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("connectionStrings");
                    this.iFLoadTableAdapter._connection.ConnectionString = connectionString;
                    Debug.WriteLine("User connectionString: " + connectionString);

                    try
                    {
                        this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                        dataSyncTimer.Start();
                    }
                    catch
                    {
                        MessageBox.Show("Could not connect to database at " + dataSource, "Database Error");
                        Application.Exit();
                    }
                }
                catch
                {
                    Properties.Settings settings = new Properties.Settings();
                    var connectionString = settings.If97VersionExpConnectionString;
                    string configDataSource = new System.Data.OleDb.OleDbConnectionStringBuilder(connectionString).DataSource;
                    string dataSource = GetDataSource(configDataSource);
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataSource;
                    //settings.If97VersionExpConnectionString = connectionString;
                    //settings.Save();
                    this.iFLoadTableAdapter._connection.ConnectionString = connectionString;
                    Debug.WriteLine("User connectionString: " + connectionString);
                    try
                    {
                        this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                        dataSyncTimer.Start();
                        using (XmlWriter writer = XmlWriter.Create("LibraInterfaceTool.exe.config"))
                        {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("configuration");
                            writer.WriteStartElement("configSections");
                            writer.WriteEndElement();
                            writer.WriteStartElement("connectionStrings");
                            writer.WriteStartElement("add");
                            writer.WriteAttributeString("name", "IF97Tool.Properties.Settings.If97VersionExpConnectionString");
                            writer.WriteAttributeString("connectionString", connectionString);
                            writer.WriteAttributeString("providerName", "System.Data.OleDb");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndDocument();
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not connect to database at " + dataSource, "Database Error");
                        Application.Exit();
                    }
                }

            }
        }

        private string GetDataSource(string configDataSource)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Database Location",
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { AutoSize = true, Left = 50, Top = 20, Text = "Enter the location of the IF97VersionExp.mdb database" };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            textBox.Text = configDataSource;
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : configDataSource;
        }

        private void dataSyncTimer_Tick(object sender, EventArgs e)
        {
            //List<DataRow> currentrows = new List<DataRow>();
            //DataGridView snapshot = new DataGridView();
            //foreach (DataGridViewColumn col in this.IFLoadTableGrid.Columns)
            //{
            //    snapshot.Columns.Add(new DataGridViewColumn(col.CellTemplate));
            //}
            //foreach (DataGridViewRow row in this.IFLoadTableGrid.Rows)
            //{
            //    snapshot.Rows.Add(row.Clone());
            //}
            //DataGridViewSelectedCellCollection selectedCells = snapshot.SelectedCells;

            try {
                int position = this.iFLoadBindingSource.Position;
                updating = true;
                this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                this.iFLoadBindingSource.Position = position;
                //IFLoadTableGrid.ClearSelection();
                IFLoadTableGrid.CurrentCell = IFLoadTableGrid.Rows[selectedRow].Cells[selectedColumn];
                if (editing)
                {
                    IFLoadTableGrid.BeginEdit(false);
                }
                if (selectedRows.Count > 0 && selectedRows != null)
                {
                    foreach(int rowSelected in selectedRows)
                    {
                        IFLoadTableGrid.Rows[rowSelected].Selected = true;
                    }
                }
                updating = false;
            }
            catch {
                updating = false;
            }

            //foreach (DataGridViewCell selectedCell in selectedCells)
            //{
            //    IFLoadTableGrid.Rows[selectedCell.RowIndex].Cells[selectedCell.ColumnIndex].Selected = true;
            //}
            //this.IFLoadTableGrid.CurrentRow.Cells[current.ColumnIndex].Selected = true;
        }

        private void IFLoadTableGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                foreach (DataGridViewRow row in IFLoadTableGrid.SelectedRows)
                {
                    IFLoadTableGrid.Rows.RemoveAt(row.Index);
                    iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                }
            }
        }

        private void IFLoadTableGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!updating)
            {
                selectedRow = IFLoadTableGrid.CurrentCell.RowIndex;
                selectedColumn = IFLoadTableGrid.CurrentCell.ColumnIndex;
                selectedRows.Clear();
                Debug.WriteLine("row=" + selectedRow + ",col=" + selectedColumn);
            }
        }

        private void deleteLoadToolstripMenuItem_Click(object sender, EventArgs e)
        {
            deleteSelectedLoads();
        }

        private void completeLoadToolstripMenuItem_Click(object sender, EventArgs e)
        {
            completeSelectedLoads();
        }

        private void IFLoadTableGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (!updating)
            {
                try {
                    iFLoadBindingSource.EndEdit();
                    int successfullyEditedRows = iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                    Debug.WriteLine("Updated " + successfullyEditedRows + " rows");
                }
                catch { }
                editing = false;
                dataSyncTimer.Start();
            }
        }

        private void IFLoadTableGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Debug.WriteLine("Edting...");
            editing = true;
            editingRow = IFLoadTableGrid.CurrentCell.RowIndex;
            editingColumn = IFLoadTableGrid.CurrentCell.ColumnIndex;
            dataSyncTimer.Stop();
        }

        private void IFLoadTableGrid_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!updating)
            {
                selectedRows.Clear();
                foreach (DataGridViewRow row in IFLoadTableGrid.SelectedRows)
                {
                    selectedRows.Add(row.Index);
                    Debug.WriteLine("Row " + row.Index + " selected");
                }
            }
        }

        private void IFLoadTableGrid_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (!updating)
            {
                selectedRows.Clear();
                foreach (DataGridViewRow row in IFLoadTableGrid.SelectedRows)
                {
                    selectedRows.Add(row.Index);
                    Debug.WriteLine("Row " + row.Index + " selected");
                }
            }
        }

        private void completeLoadBtn_Click(object sender, EventArgs e)
        {
            completeSelectedLoads();
        }

        private void deleteLoadBtn_Click(object sender, EventArgs e)
        {
            deleteSelectedLoads();
        }

        private void deleteSelectedLoads()
        {
            try
            {
                foreach (DataGridViewCell cell in IFLoadTableGrid.SelectedCells)
                {
                    IFLoadTableGrid.Rows.RemoveAt(cell.RowIndex);
                    iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                }
            }
            catch { }
        }

        private void completeSelectedLoads()
        {
            try
            {
                editing = true;
                dataSyncTimer.Stop();
                foreach (DataGridViewCell cell in IFLoadTableGrid.SelectedCells)
                {
                    //(If97VersionExpDataSet.IFLoadRow)(IFLoadTableGrid.Rows[cell.RowIndex])
                    if (cell.OwningColumn == loadStatusDataGridViewTextBoxColumn)
                    {
                        cell.Value = 3;
                    }
                }
                iFLoadBindingSource.EndEdit();
                iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                editing = false;
                dataSyncTimer.Start();
            }
            catch { }
        }


    }
}
