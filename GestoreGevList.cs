using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    class GestoreGevList
    {

        public DataTable getGevListFromDB(Form1 formOne) {
            String queryString;
            DataTable dTable = new DataTable();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try {
                connection.Open();
            }
            catch (OleDbException ode) {
                MessageBox.Show("ATTENZIONE: " + ode);
            }

            queryString = "SELECT Gev_ID, Radio, Cognome, Nome, [e-mail] as email, [Telefono Fisso] as fisso, Ruolo ";
            queryString = queryString + "FROM tab_DatiAnagrafici WHERE ";
            queryString = queryString + "Cognome LIKE '" + formOne.textBox3.Text + "%' ";
            queryString = queryString + "AND Nome LIKE '" + formOne.textBox4.Text + "%' ";
            if (!formOne.textBox5.Text.Equals(""))
              queryString = queryString + "AND Radio LIKE '%" + formOne.textBox5.Text + "%' ";
            if (!formOne.textBox6.Text.Equals(""))
              queryString = queryString + "AND [e-mail] LIKE '%" + formOne.textBox6.Text + "%' ";
            if (!formOne.textBox7.Text.Equals(""))
              queryString = queryString + "AND [Telefono Fisso] LIKE '%" + formOne.textBox7.Text + "%' ";
            if (formOne.comboBox5.SelectedIndex > 0)
                queryString = queryString + "AND Ruolo = '" + formOne.comboBox5.SelectedItem.ToString() + "' ";
            queryString = queryString + "ORDER BY Cognome ";

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dTable.Clear();
            dataAdapter.Fill(dTable);

            connection.Close();
            connection.Dispose();

            return dTable;
        }

        public void getRuoliList(Form1 formOne)
        {
            String queryString;
            DataTable dtResult = new DataTable();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }
            queryString = "SELECT Codice ";
            queryString = queryString + "FROM tabd_Ruoli ";

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dtResult.Clear();
            dataAdapter.Fill(dtResult);

            connection.Close();
            connection.Dispose();
//            formOne.comboBox5.Items.Clear();
            formOne.comboBox5.Items.Add("Tutti ...");
            foreach (DataRow riga in dtResult.Rows) {
                formOne.comboBox5.Items.Add(riga["Codice"]);
            }
        }

        public void designGevList(Form1 formOne)
        {
            formOne.MaximumSize = new System.Drawing.Size(920, 720);
            formOne.MinimumSize = new System.Drawing.Size(920, 720);
            
            formOne.label3.Text = "Lista dati anagrafici GEV";
            formOne.label3.ForeColor = System.Drawing.Color.MediumSeaGreen;

            // Rimuovi il pannello del filtro ods
            formOne.Controls.Remove(formOne.panel3);
            // Rimuovi il pannello con la gridView1 degli ods
            formOne.Controls.Remove(formOne.panel4);
            // Rimuovi il pannello del dettaglio dell'ods
            formOne.Controls.Remove(formOne.panel7);
            // Rimuovi il pannello del dettaglio dati gev
            formOne.Controls.Remove(formOne.panel8);

            // Aggiungi il pannello con la gridView2 dell'anagrafica
            formOne.Controls.Add(formOne.panel6);
        }

        public int retrieveGevID(DataGridViewCellEventArgs e, DataGridView dataGridView)
        {
            object value = dataGridView.Rows[e.RowIndex].Cells["gevid"].Value;
            //if (value == null) {
            //    MessageBox.Show("WARNING: gev id non trovato");
            //    return -1;
            //}
            //return (int)value;
            return value == null ? -1 : (int)value;
        }
    }
}
