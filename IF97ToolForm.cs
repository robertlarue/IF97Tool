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
using System.IO;

namespace IF97Tool
{
    public partial class IF97ToolForm : Form
    {
        public static int editingRowLoads = 0;
        public static int editingColumnLoads = 0;
        public static int selectedRowLoads = 0;
        public static int selectedColumnLoads = 0;
        public static List<int> selectedRowsLoads = new List<int>();
        public static bool editingLoads = false;
        public static bool updatingLoads = false;
        public static int scrollPositionLoads = 0;

        public static int editingRowQueue = 0;
        public static int editingColumnQueue = 0;
        public static int selectedRowQueue = 0;
        public static int selectedColumnQueue = 0;
        public static List<int> selectedRowsQueue = new List<int>();
        public static bool editingQueue = false;
        public static bool updatingQueue = false;
        public static int scrollPositionQueue = 0;

        public IF97ToolForm()
        {
            
            InitializeComponent();
        }

        private void IF97ToolForm_Load(object sender, EventArgs e)
        {
            if (!File.Exists(System.AppDomain.CurrentDomain.FriendlyName + ".Config"))
            {
                File.WriteAllText(System.AppDomain.CurrentDomain.FriendlyName + ".Config", Properties.Resources.app_config);
            }
        }

        private void IF97ToolForm_Shown(object sender, EventArgs e)
        {

            try
            {
                var appSettings = new Properties.Settings();
                this.iFLoadTableAdapter._connection.ConnectionString = appSettings.If97VersionExpConnectionString;
                this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                this.queueTrucksTableAdapter._connection.ConnectionString = appSettings.Gen3DataConnectionString;
                this.queueTrucksTableAdapter.Fill(this.gen3DataDataSet.QueueTrucks);
                dataSyncTimer.Start();
            }
            catch
            {
                try
                {
                    this.iFLoadTableAdapter._connection.ConnectionString = Properties.Settings.Default.If97VersionExpConnectionString;
                    this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                    this.queueTrucksTableAdapter._connection.ConnectionString = Properties.Settings.Default.Gen3DataConnectionString;
                    this.queueTrucksTableAdapter.Fill(this.gen3DataDataSet.QueueTrucks);
                    dataSyncTimer.Start();
                }
                catch
                {
                    try
                    {
                        this.iFLoadTableAdapter._connection.ConnectionString = Properties.Settings.Default.If97VersionExpConnectionStringServer;
                        this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                        this.queueTrucksTableAdapter._connection.ConnectionString = Properties.Settings.Default.Gen3DataConnectionStringServer;
                        this.queueTrucksTableAdapter.Fill(this.gen3DataDataSet.QueueTrucks);
                        dataSyncTimer.Start();
                    }
                    catch
                    {
                        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                        var connectionStringsSection = (ConnectionStringsSection)config.GetSection("connectionStrings");
                        var connectionStringLoads = connectionStringsSection.ConnectionStrings["IF97Tool.Properties.Settings.If97VersionExpConnectionString"].ConnectionString;
                        var connectionStringQueue = connectionStringsSection.ConnectionStrings["IF97Tool.Properties.Settings.Gen3DataConnectionString"].ConnectionString;
                        Debug.WriteLine("Original connectionStringLoads: " + connectionStringLoads);
                        Debug.WriteLine("Original connectionStringQueue: " + connectionStringQueue);

                        string configDataSourceLoads = new System.Data.OleDb.OleDbConnectionStringBuilder(connectionStringLoads).DataSource;
                        string dataSourceLoads = GetDataSource(configDataSourceLoads);
                        connectionStringLoads = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataSourceLoads;
                        connectionStringsSection.ConnectionStrings["IF97Tool.Properties.Settings.If97VersionExpConnectionString"].ConnectionString = connectionStringLoads;

                        string configDataSourceQueue = new System.Data.OleDb.OleDbConnectionStringBuilder(connectionStringQueue).DataSource;
                        string dataSourceQueue = GetDataSource(configDataSourceQueue);
                        connectionStringQueue = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataSourceQueue;
                        connectionStringsSection.ConnectionStrings["IF97Tool.Properties.Settings.Gen3DataConnectionString"].ConnectionString = connectionStringQueue;

                        config.Save(ConfigurationSaveMode.Modified);
                        ConfigurationManager.RefreshSection("connectionStrings");
                        this.iFLoadTableAdapter._connection.ConnectionString = connectionStringLoads;
                        this.queueTrucksTableAdapter._connection.ConnectionString = connectionStringQueue;
                        Debug.WriteLine("User connectionStringLoads: " + connectionStringLoads);
                        Debug.WriteLine("User connectionStringQueue: " + connectionStringQueue);

                        try
                        {
                            this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                        }
                        catch
                        {
                            MessageBox.Show("Could not connect to database at " + dataSourceLoads, "Database Error");
                            Application.Exit();
                        }
                        try
                        {
                            this.queueTrucksTableAdapter.Fill(this.gen3DataDataSet.QueueTrucks);
                        }
                        catch
                        {
                            MessageBox.Show("Could not connect to database at " + dataSourceQueue, "Database Error");
                            Application.Exit();
                        }
                        dataSyncTimer.Start();
                    }
                }
            }
        }

        private string GetDataSource(string configDataSource)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 175,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Database Location",
                StartPosition = FormStartPosition.CenterScreen
            };
            string databaseFile = Path.GetFileName(configDataSource);
            Label textLabel = new Label() { AutoSize = true, Left = 50, Top = 20, Text = "Enter the location of the " + databaseFile + " database" };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            textBox.Text = configDataSource;
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Height = 30, Top = 90, DialogResult = DialogResult.OK };
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
                updatingLoads = true;
                this.iFLoadTableAdapter.Fill(this.if97VersionExpDataSet.IFLoad);
                this.iFLoadBindingSource.Position = position;
                //IFLoadTableGrid.ClearSelection();
                IFLoadTableGrid.CurrentCell = IFLoadTableGrid.Rows[selectedRowLoads].Cells[selectedColumnLoads];
                if (editingLoads)
                {
                    IFLoadTableGrid.BeginEdit(false);
                }
                if (selectedRowsLoads.Count > 0 && selectedRowsLoads != null)
                {
                    foreach(int rowSelected in selectedRowsLoads)
                    {
                        IFLoadTableGrid.Rows[rowSelected].Selected = true;
                    }
                }
                IFLoadTableGrid.HorizontalScrollingOffset = scrollPositionLoads;
                updatingLoads = false;
            }
            catch {
                updatingLoads = false;
            }

            try
            {
                int position = this.gen3DataDataSetBindingSource.Position;
                updatingQueue = true;
                this.queueTrucksTableAdapter.Fill(this.gen3DataDataSet.QueueTrucks);
                this.queueTrucksBindingSource.Position = position;
                queueTrucksTableGrid.CurrentCell = queueTrucksTableGrid.Rows[selectedRowQueue].Cells[selectedColumnQueue];
                if (editingQueue)
                {
                    queueTrucksTableGrid.BeginEdit(false);
                }
                if (selectedRowsQueue.Count > 0 && selectedRowsQueue != null)
                {
                    foreach (int rowSelected in selectedRowsQueue)
                    {
                        queueTrucksTableGrid.Rows[rowSelected].Selected = true;
                    }
                }
                queueTrucksTableGrid.HorizontalScrollingOffset = scrollPositionQueue;
                updatingQueue = false;
            }
            catch
            {
                updatingQueue = false;
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
                deleteSelectedLoads();
            }
        }

        private void IFLoadTableGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!updatingLoads)
            {
                selectedRowLoads = IFLoadTableGrid.CurrentCell.RowIndex;
                selectedColumnLoads = IFLoadTableGrid.CurrentCell.ColumnIndex;
                selectedRowsLoads.Clear();
                Debug.WriteLine("row=" + selectedRowLoads + ",col=" + selectedColumnLoads);
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
            if (!updatingLoads)
            {
                try {
                    iFLoadBindingSource.EndEdit();
                    int successfullyEditedRows = iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                    Debug.WriteLine("Updated " + successfullyEditedRows + " rows");
                }
                catch { }
                editingLoads = false;
                dataSyncTimer.Start();
            }
        }

        private void IFLoadTableGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Debug.WriteLine("Edting...");
            editingLoads = true;
            editingRowLoads = IFLoadTableGrid.CurrentCell.RowIndex;
            editingColumnLoads = IFLoadTableGrid.CurrentCell.ColumnIndex;
            dataSyncTimer.Stop();
        }

        private void IFLoadTableGrid_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!updatingLoads)
            {
                selectedRowsLoads.Clear();
                foreach (DataGridViewRow row in IFLoadTableGrid.SelectedRows)
                {
                    selectedRowsLoads.Add(row.Index);
                    Debug.WriteLine("Row " + row.Index + " selected");
                }
            }
        }

        private void IFLoadTableGrid_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (!updatingLoads)
            {
                selectedRowsLoads.Clear();
                foreach (DataGridViewRow row in IFLoadTableGrid.SelectedRows)
                {
                    selectedRowsLoads.Add(row.Index);
                    Debug.WriteLine("Row " + row.Index + " selected");
                }
            }
        }

        private void completeLoadBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("This may create an error ticket. Do you want to continue?", "Complete Load", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                completeSelectedLoads();
            }
        }

        private void deleteLoadBtn_Click(object sender, EventArgs e)
        {
            deleteSelectedLoads();
        }

        private void deleteSelectedLoads()
        {
            try
            {
                if (IFLoadTableGrid.SelectedRows.Count == 0)
                {
                    foreach (DataGridViewCell cell in IFLoadTableGrid.SelectedCells)
                    {
                        IFLoadTableGrid.Rows.RemoveAt(cell.RowIndex);
                        iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                    }
                }
                else
                {
                    foreach (DataGridViewRow row in IFLoadTableGrid.SelectedRows)
                    {
                        IFLoadTableGrid.Rows.RemoveAt(row.Index);
                        iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                    }
                }
            }
            catch { }
        }

        private void deleteSelectedQueueTruck()
        {
            try
            {
                if (queueTrucksTableGrid.SelectedRows.Count == 0)
                {
                    foreach (DataGridViewCell cell in queueTrucksTableGrid.SelectedCells)
                    {
                        queueTrucksTableGrid.Rows.RemoveAt(cell.RowIndex);
                        queueTrucksTableAdapter.Update(gen3DataDataSet.QueueTrucks);
                    }
                }
                else
                {
                    foreach (DataGridViewRow row in queueTrucksTableGrid.SelectedRows)
                    {
                        queueTrucksTableGrid.Rows.RemoveAt(row.Index);
                        queueTrucksTableAdapter.Update(gen3DataDataSet.QueueTrucks);
                    }
                }
            }
            catch { }
        }

        private void completeSelectedLoads()
        {
            try
            {
                editingLoads = true;
                dataSyncTimer.Stop();
                if (queueTrucksTableGrid.SelectedRows.Count == 0)
                {
                    foreach (DataGridViewCell cell in IFLoadTableGrid.SelectedCells)
                    {
                        cell.OwningRow.Cells[5].Value = 3;
                    }
                }
                else
                {
                    foreach (DataGridViewCell cell in IFLoadTableGrid.SelectedCells)
                    {
                        if (cell.OwningColumn == loadStatusDataGridViewTextBoxColumn)
                        {
                            cell.Value = 3;
                        }
                    }
                }
                iFLoadBindingSource.EndEdit();
                iFLoadTableAdapter.Update(if97VersionExpDataSet.IFLoad);
                editingLoads = false;
                dataSyncTimer.Start();
            }
            catch { }
        }

        private void queueDeleteButton_Click(object sender, EventArgs e)
        {
            deleteSelectedQueueTruck();
        }

        private void queueTrucksTableGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Debug.WriteLine("Edting...");
            editingQueue = true;
            editingRowQueue = queueTrucksTableGrid.CurrentCell.RowIndex;
            editingColumnQueue = queueTrucksTableGrid.CurrentCell.ColumnIndex;
            dataSyncTimer.Stop();
        }

        private void queueTrucksTableGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (!updatingQueue)
            {
                try
                {
                    queueTrucksBindingSource.EndEdit();
                    int successfullyEditedRows = queueTrucksTableAdapter.Update(gen3DataDataSet.QueueTrucks);
                    Debug.WriteLine("Updated " + successfullyEditedRows + " rows");
                }
                catch { }
                editingQueue = false;
                dataSyncTimer.Start();
            }
        }

        private void queueTrucksTableGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!updatingQueue)
            {
                selectedRowQueue = queueTrucksTableGrid.CurrentCell.RowIndex;
                selectedColumnQueue = queueTrucksTableGrid.CurrentCell.ColumnIndex;
                selectedRowsQueue.Clear();
                Debug.WriteLine("row=" + selectedRowQueue + ",col=" + selectedColumnQueue);
            }
        }

        private void queueTrucksTableGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                deleteSelectedQueueTruck();
            }
        }

        private void queueTrucksTableGrid_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (!updatingQueue)
            {
                selectedRowsQueue.Clear();
                foreach (DataGridViewRow row in queueTrucksTableGrid.SelectedRows)
                {
                    selectedRowsQueue.Add(row.Index);
                    Debug.WriteLine("Row " + row.Index + " selected");
                }
            }
        }

        private void queueTrucksTableGrid_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!updatingQueue)
            {
                selectedRowsQueue.Clear();
                foreach (DataGridViewRow row in queueTrucksTableGrid.SelectedRows)
                {
                    selectedRowsQueue.Add(row.Index);
                    Debug.WriteLine("Row " + row.Index + " selected");
                }
            }
        }

        private void IFLoadTableGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (!updatingLoads)
            {
                scrollPositionLoads = e.NewValue;
            }
        }

        private void queueTrucksTableGrid_Scroll(object sender, ScrollEventArgs e)
        {
            if (!updatingQueue)
            {
                scrollPositionQueue = e.NewValue;
            }
        }
    }
}
