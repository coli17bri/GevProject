using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public partial class FormSelezionePartecipanti : Form
    {
        private String odsId;

        public void setOdsId (String pOdsId)
        { odsId = pOdsId; }

        public String getOdsId()
        { return odsId; }

        public FormSelezionePartecipanti()
        {
            InitializeComponent();
        }

        private void FormSelezionePartecipanti_Load(object sender, EventArgs e)
        {
            String queryString;
            DataTable dTable = new DataTable();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT gev_ID, Cognome, Nome, Radio ";
            queryString = queryString + "FROM tab_DatiAnagrafici INNER JOIN tabd_Ruoli ";
            queryString = queryString + "ON tab_DatiAnagrafici.ruolo = tabd_Ruoli.Codice ";
            queryString = queryString + "WHERE NOT tabd_Ruoli.ToIgnore ";
            queryString = queryString + "AND Gev_ID NOT IN (";
            queryString = queryString + "SELECT Gev_ID FROM tab_rel_Gev_Ods WHERE Ods_ID = " + this.getOdsId() + " ";
            queryString = queryString + "AND tab_rel_Gev_Ods.D_END = cDate('31/12/9999')) ";
            queryString = queryString + "ORDER BY Cognome ";

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dTable.Clear();
            dataAdapter.Fill(dTable);

            connection.Close();
            connection.Dispose();


            DataGridViewCheckBoxColumn selectedDataGridViewColumn = new DataGridViewCheckBoxColumn();
            DataGridViewTextBoxColumn gevIdDataGridViewColumn = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn cognomeDataGridViewColumn = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn nomeDataGridViewColumn = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn radioDataGridViewColumn = new DataGridViewTextBoxColumn();

            selectedDataGridViewColumn.HeaderText = "";
            selectedDataGridViewColumn.Width = 50;
            selectedDataGridViewColumn.Name = "selection";
            selectedDataGridViewColumn.ReadOnly = false;

            gevIdDataGridViewColumn.DataPropertyName = "gev_ID";
            gevIdDataGridViewColumn.HeaderText = "Gev_ID";
            gevIdDataGridViewColumn.Name = "gevid";
            gevIdDataGridViewColumn.ReadOnly = true;
            gevIdDataGridViewColumn.Visible = false;

            cognomeDataGridViewColumn.DataPropertyName = "Cognome";
            cognomeDataGridViewColumn.HeaderText = "Cognome";
            cognomeDataGridViewColumn.ReadOnly = true;
            cognomeDataGridViewColumn.Width = 160;
            cognomeDataGridViewColumn.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 10F); 

            nomeDataGridViewColumn.DataPropertyName = "Nome";
            nomeDataGridViewColumn.HeaderText = "Nome";
            nomeDataGridViewColumn.Width = 130;
            nomeDataGridViewColumn.ReadOnly = true;

            radioDataGridViewColumn.DataPropertyName = "Radio";
            radioDataGridViewColumn.HeaderText = "N. radio";
            radioDataGridViewColumn.Width = 60;
            radioDataGridViewColumn.ReadOnly = true;
            radioDataGridViewColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


            this.dataGridView1.Columns.Clear();
            this.dataGridView1.Columns.AddRange(new DataGridViewColumn[] {
                selectedDataGridViewColumn,
                gevIdDataGridViewColumn,
                cognomeDataGridViewColumn,
                nomeDataGridViewColumn,
                radioDataGridViewColumn
                });


            this.dataGridView1.DataSource = dTable;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String strSQL;
            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }
            // L'adapter serve solo se mi viene passato un dataTable o dataSet da aggiornare
            // OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
            OleDbCommand dbCommand = new OleDbCommand();

            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                DataGridViewCell cell = row.Cells["selection"];
                if (cell != null && cell.Value != null && (Boolean)cell.Value)
                {                    
                  strSQL = "INSERT INTO tab_rel_Gev_Ods([Gev_ID], [Ods_ID]) ";
                  strSQL = strSQL + "VALUES (" + row.Cells["gevid"].Value + "," + this.getOdsId() + ");";
                  dbCommand.CommandText = strSQL;
                  dbCommand.Connection = connection;

                  dbCommand.ExecuteNonQuery();
                }
            }

            connection.Close();
            connection.Dispose();
            this.DialogResult = DialogResult.OK;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
