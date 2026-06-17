using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    class GestoreOdsList
    {

        public DataTable getOdsListFromDB(FiltroOds filtro) {
            String queryString;
            DataTable dTable = new DataTable();
            String resultPeople = "Qualcuno...";

            OleDbConnection connection = ConnectionHelper.createConnection();
            try {
                connection.Open();
            }
            catch (OleDbException ode) {
              MessageBox.Show("ATTENZIONE: " + ode);  
            }

            queryString = "SELECT Ods_id, Numero, Data, Durata, [CODICE SERVIZIO] as tipoServizio, [CODICE ZONA] as luogo FROM tab_OrdiniDiServizio ";
            queryString = queryString + "WHERE Numero BETWEEN '" + filtro.getNumeroFrom() + "' AND '" + filtro.getNumeroTo() + "' ";
            queryString = queryString + "AND  Data BETWEEN cDate('" + filtro.getDataFrom().ToShortDateString() + "') AND cDate('" + filtro.getDataTo().ToShortDateString() + "') ";
            if (!filtro.getTipo().Equals("-1"))
                queryString = queryString + "AND [CODICE SERVIZIO] = '" + filtro.getTipo() + "' ";
            if (!filtro.getLuogo().Equals("-1"))
            {
              if (filtro.getLuogo().Equals(""))
                  queryString = queryString + "AND (IsNull([CODICE ZONA]) OR [CODICE ZONA] = '') ";
              else
                  queryString = queryString + "AND [CODICE ZONA] = '" + filtro.getLuogo() + "' ";
            }
            queryString = queryString + "AND D_END = cDate('31/12/9999') ";
            queryString = queryString + "ORDER BY Data DESC";

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dTable.Clear();
            dataAdapter.Fill(dTable);
            dTable.Columns.Add("partecipanti", typeof(String));

            foreach (DataRow riga in dTable.Rows)
            {
                resultPeople = "";
                queryString = "SELECT tab_DatiAnagrafici.Cognome ";
                queryString = queryString + "FROM tab_DatiAnagrafici INNER JOIN tab_rel_Gev_Ods ON tab_DatiAnagrafici.Gev_ID = tab_rel_Gev_Ods.Gev_ID ";
                queryString = queryString + "WHERE  tab_rel_Gev_Ods.ODS_ID = " + riga["Ods_id"] + " ";
                queryString = queryString + "AND tab_rel_Gev_Ods.D_END = cDate('31/12/9999') ";
                queryString = queryString + "ORDER BY Cognome ";

                OleDbCommand dbCommand = new OleDbCommand(queryString, connection);
                OleDbDataReader dataReader = dbCommand.ExecuteReader();
                while (dataReader.Read()) 
                    resultPeople += dataReader["Cognome"] + "; ";
                dataReader.Close();

                riga["partecipanti"] = resultPeople;
            }

            connection.Close();
            connection.Dispose();

            return dTable;
        }

        public void fillComboFilter(Form1 formOne) {
            String queryString;

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT Codice FROM tabd_Servizi";
            OleDbCommand dbCommand = new OleDbCommand(queryString, connection);
            OleDbDataReader dataReader = dbCommand.ExecuteReader();
            formOne.comboBox1.Items.Add("Select code ...");
            formOne.comboBox4.Items.Add("Select code ...");
            while (dataReader.Read())
            {
                formOne.comboBox1.Items.Add(dataReader["Codice"]);
                formOne.comboBox4.Items.Add(dataReader["Codice"]);
            }
            formOne.comboBox1.SelectedIndex = 0;
            formOne.comboBox4.SelectedIndex = 0;

            queryString = "SELECT Codice FROM tabd_Zone";
            dbCommand = new OleDbCommand(queryString, connection);
            dataReader = dbCommand.ExecuteReader();
            formOne.comboBox3.Items.Add("Select ...");
            while (dataReader.Read())
            {
                formOne.comboBox3.Items.Add(dataReader["Codice"]);
            }
            formOne.comboBox3.Items.Add("");
            formOne.comboBox3.SelectedIndex = 0;

            connection.Close();
            connection.Dispose();
        }

        public void designOdsList(Form1 formOne)
        {
            formOne.MaximumSize = new System.Drawing.Size(920, 720);
            formOne.MinimumSize = new System.Drawing.Size(920, 720);

            formOne.label3.Text = "Lista ordini di servizio";
            formOne.label3.ForeColor = System.Drawing.Color.DodgerBlue;

            // Rimuovi il pannello con la gridView2 dell'anagrafica
            formOne.Controls.Remove(formOne.panel6);
            // Rimuovi il pannello del dettaglio dell'ods
            formOne.Controls.Remove(formOne.panel7);
            // Rimuovi il pannello del dettaglio dati gev
            formOne.Controls.Remove(formOne.panel8);

            // Aggiungi il pannello del filtro ods
            formOne.Controls.Add(formOne.panel3);
            // Aggiungi il pannello con la gridView1 degli ods
            formOne.Controls.Add(formOne.panel4);
        }

        public int retrieveOdsID(DataGridViewCellEventArgs e, DataGridView dataGridView) {
            object value = dataGridView.Rows[e.RowIndex].Cells["odsid"].Value;
            if (value is DBNull) {
                MessageBox.Show("WARNING: ods id non trovato");
                return -1; 
            }
            return (int) value;
        }

    }
}
