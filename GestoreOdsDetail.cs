using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Drawing;

namespace WindowsFormsApplication1
{
    class GestoreOdsDetail
    {

        public DataRow getOdsDetailFromDB (int odsKey) {
            String queryString;
            DataSet dsResult = new DataSet();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try {
                connection.Open();
            }
            catch (OleDbException ode) {
              MessageBox.Show("ATTENZIONE: " + ode);  
            }
            queryString = "SELECT ODS_ID, Numero, Data, DURATA, [CODICE SERVIZIO], [CODICE ZONA], Commenti, Consegnato, Approvato ";
            queryString = queryString + "FROM tab_OrdiniDiServizio ";
            queryString = queryString + "WHERE tab_OrdiniDiServizio.ODS_ID=" + odsKey;

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dataAdapter.Fill(dsResult, "DetailODS");

            connection.Close();
            connection.Dispose();
            return dsResult.Tables["DetailODS"].Rows[0];
        }

        public int searchOds(String number)
        {
            number = GeneralHelper.formatNumOds(number);
            OleDbConnection connection = ConnectionHelper.createConnection();

            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            OleDbCommand cmd = 
                  new OleDbCommand("SELECT Ods_id FROM tab_OrdiniDiServizio WHERE Numero = '" + number + "' " +
                                    "AND D_END = cDate('31/12/9999') " , connection);
            OleDbDataReader dataReader = cmd.ExecuteReader();
            if (dataReader.HasRows)
            {
                dataReader.Read();
                if (!dataReader.IsDBNull(0))
                    return Convert.ToInt32(dataReader["Ods_id"]);
            }
            dataReader.Close();

            connection.Close();
            connection.Dispose();

            return -1;
        }


        public DataTable getOdsPartecipanti(int odskey) {
            String queryString;
            DataTable dTable = new DataTable();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT tab_rel_Gev_Ods.GEV_ID AS gev_ID, ";
            queryString += "tab_rel_Gev_Ods.ODS_ID AS ods_ID, ";
            queryString += "tab_DatiAnagrafici.Cognome AS Cognome, tab_DatiAnagrafici.Nome AS Nome, ";
            queryString += "tab_DatiAnagrafici.Radio AS [Numero radio] ";
            queryString += "FROM tab_DatiAnagrafici INNER JOIN tab_rel_Gev_Ods ON tab_DatiAnagrafici.Gev_ID = tab_rel_Gev_Ods.Gev_ID ";
            queryString += "WHERE  tab_rel_Gev_Ods.ODS_ID = " + odskey + " ";
            queryString += "AND tab_rel_Gev_Ods.D_END = cDate('31/12/9999') ";
            queryString += "ORDER BY Cognome ";

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dTable.Clear();
            dataAdapter.Fill(dTable);

            connection.Close();
            connection.Dispose();

            return dTable;
        }

        public DataSet getComboxList()
        {
            String queryString;
            DataSet dsResult = new DataSet();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }
            queryString = "SELECT Codice, Descrizione ";
            queryString = queryString + "FROM tabd_Servizi ";
            queryString = queryString + "WHERE Descrizione <> '' ";
            queryString = queryString + "ORDER BY Codice";

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dataAdapter.Fill(dsResult, "ListaServizi");

            queryString = "SELECT Codice, Descrizione ";
            queryString = queryString + "FROM tabd_Zone ";
            //queryString = queryString + "WHERE Descrizione <> ''";
            dataAdapter = new OleDbDataAdapter(queryString, connection);
            dataAdapter.Fill(dsResult, "ListaZone");

            connection.Close();
            connection.Dispose();
            return dsResult;
        }

        // DELETE 
        public int deleteOdsDetail(int odsId)
        {
            String strSQL;
            int rowsAffected = 0;
            OleDbConnection connection = ConnectionHelper.createConnection();
            OleDbCommand dbCommand = null;
            bool isOperationOK = true;

            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            try
            {
                strSQL = "UPDATE tab_rel_Gev_Ods ";
                strSQL = strSQL + "SET D_END = Date() ";
                strSQL = strSQL + "WHERE Ods_ID = " + odsId;
                dbCommand = new OleDbCommand(strSQL, connection);
                rowsAffected = dbCommand.ExecuteNonQuery();
            }
            catch (System.InvalidOperationException ioe) {
                isOperationOK = false;
                MessageBox.Show("Delete ods relation exception: " + ioe);
            }

            if (isOperationOK) {
                try {
                    strSQL = "UPDATE tab_OrdiniDiServizio ";
                    strSQL = strSQL + "SET D_END = Date() ";
                    strSQL = strSQL + "WHERE Ods_ID = " + odsId;

                    dbCommand = new OleDbCommand(strSQL, connection);
                    dbCommand.ExecuteNonQuery();
                }
                catch (System.InvalidOperationException ioe) {
                   MessageBox.Show("Delete ods exception: " + ioe);
                }
            }

            connection.Close();
            connection.Dispose();

            return rowsAffected;
        }

        public int deleteOdsPartecipante(int gevId, int odsId)
        {
            String strSQL;
            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            //strSQL = "DELETE * FROM tab_rel_Gev_Ods ";
            //strSQL = strSQL + "WHERE Gev_ID = " + gevId + " AND Ods_ID = " + odsId ;
            strSQL = "UPDATE tab_rel_Gev_Ods ";
            strSQL = strSQL + "SET D_END = Date() " ;
            strSQL = strSQL + "WHERE Gev_ID = " + gevId + " AND Ods_ID = " + odsId;

            OleDbCommand dbCommand = new OleDbCommand(strSQL, connection);
            int rowsAffected = dbCommand.ExecuteNonQuery();

            connection.Close();
            connection.Dispose();

            return rowsAffected;
        }

        public int insertOdsDetail(Form1 formOne)
        {
            String strSQL;
            OleDbConnection connection = ConnectionHelper.createConnection();

            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            strSQL = "INSERT INTO tab_OrdiniDiServizio(Numero, Data, DURATA, [CODICE SERVIZIO], [CODICE ZONA], Commenti, Approvato, Consegnato) ";
            strSQL = strSQL + "VALUES ('" + GeneralHelper.formatNumOds(formOne.textBox_odsNumero.Text) + "', "
                                          + "cDate('" + formOne.dateTimePicker_odsEffettuato.Value + "'), "
                                          + "cDate('" + GeneralHelper.formatDurataOds(formOne.textBox_odsDurata.Text) + "'), "
                                          + "'" + formOne.comboBox_odsTipo.SelectedValue  + "', "
                                          + "'" + formOne.comboBox_odsLocalita.SelectedValue + "', "
                                          + "?, "
                                          //+ "'" + formOne.richTextBox1.Text + "', "
                                          + formOne.checkBox1.Checked + ", "
                                          + formOne.checkBox2.Checked 
                                          + ");";
            OleDbCommand dbCommand = new OleDbCommand(strSQL, connection);
            dbCommand.Parameters.Clear();
            OleDbParameter paramCommenti = new OleDbParameter("@pComm", OleDbType.VarChar, 300);
            dbCommand.Parameters.Add(paramCommenti).Value = formOne.richTextBox1.Text;
            int rowsAffected = dbCommand.ExecuteNonQuery();

            OleDbCommand cmd = new OleDbCommand("Select max(Ods_id) as lastOdsId from tab_OrdiniDiServizio", connection);
            OleDbDataReader dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
                formOne.label_odsId.Text = dataReader["lastOdsId"].ToString();
            dataReader.Close();

            connection.Close();
            connection.Dispose();

            return rowsAffected;
        }

        public int updateOdsDetail(Form1 formOne, int odsId)
        {
            String strSQL;
            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            strSQL = "UPDATE tab_OrdiniDiServizio ";
            strSQL = strSQL + "SET Numero = '" + GeneralHelper.formatNumOds(formOne.textBox_odsNumero.Text) + "', "
                            + "Data = cDate('" + formOne.dateTimePicker_odsEffettuato.Value + "'), "
                            + "DURATA = cDate('" + formOne.textBox_odsDurata.Text + "'), "
                            + "[CODICE SERVIZIO] = '" + formOne.comboBox_odsTipo.SelectedValue + "', "
                            + "[CODICE ZONA] = '" + formOne.comboBox_odsLocalita.SelectedValue + "', "
                            + "Commenti = ?, "
                            // + "Commenti = '" + formOne.richTextBox1.Text + "', "
                            + "Approvato = " + formOne.checkBox1.Checked + ", "
                            + "Consegnato = " + formOne.checkBox2.Checked + " ";
            strSQL = strSQL + "WHERE Ods_ID = " + odsId;
                            
            OleDbCommand dbCommand = new OleDbCommand(strSQL, connection);
            dbCommand.Parameters.Clear();
            OleDbParameter paramCommenti = new OleDbParameter("@pComm", OleDbType.VarChar, 300);
            dbCommand.Parameters.Add(paramCommenti).Value = formOne.richTextBox1.Text;
            int rowsAffected = dbCommand.ExecuteNonQuery();

            connection.Close();
            connection.Dispose();

            return rowsAffected;
        }

        public void fillCombobox(Form1 formOne)
        {
            DataSet dataSetCombo = this.getComboxList();

            formOne.comboBox_odsTipo.DataSource = dataSetCombo.Tables["ListaServizi"];
            formOne.comboBox_odsTipo.DisplayMember = "Descrizione";
            formOne.comboBox_odsTipo.ValueMember = "Codice";
            GeneralHelper.createMulticolumnCombobox(formOne.comboBox_odsTipo);

            formOne.comboBox_odsLocalita.DataSource = dataSetCombo.Tables["ListaZone"];
            formOne.comboBox_odsLocalita.DisplayMember = "Descrizione";
            formOne.comboBox_odsLocalita.ValueMember = "Codice";
            GeneralHelper.createMulticolumnCombobox(formOne.comboBox_odsLocalita, 0.1);        
        }

        public void fillOdsDetailForm(Form1 formOne, int odsIdSelected)
        {
            this.fillCombobox(formOne);
            this.designOdsDetail(formOne);

            DataRow datiBaseOds = this.getOdsDetailFromDB(odsIdSelected);
            formOne.label_odsId.Text = datiBaseOds.ItemArray.GetValue(0).ToString();
            formOne.textBox_odsNumero.Text = datiBaseOds.ItemArray.GetValue(1).ToString();
            formOne.dateTimePicker_odsEffettuato.Value = (DateTime)datiBaseOds.ItemArray.GetValue(2);
            formOne.textBox_odsDurata.Text = ((DateTime)datiBaseOds.ItemArray.GetValue(3)).ToShortTimeString();
            formOne.label31.Text = ((DataRowView)formOne.comboBox_odsTipo.Items[0]).Row["Codice"].ToString();
            formOne.comboBox_odsTipo.SelectedValue = datiBaseOds.ItemArray.GetValue(4).ToString();
            formOne.label32.Text = ((DataRowView)formOne.comboBox_odsLocalita.Items[0]).Row["Codice"].ToString();
            formOne.comboBox_odsLocalita.SelectedValue = datiBaseOds.ItemArray.GetValue(5).ToString();
            formOne.richTextBox1.Text = datiBaseOds.ItemArray.GetValue(6).ToString();
            formOne.checkBox2.Checked = (bool)datiBaseOds.ItemArray.GetValue(7);
            formOne.checkBox1.Checked = (bool)datiBaseOds.ItemArray.GetValue(8);

            formOne.dataGridView3.DataSource = this.getOdsPartecipanti(odsIdSelected);
            formOne.dataGridView3.ClearSelection();

            formOne.dataGridView3.Visible = true;
            formOne.label15.Visible = true;
            formOne.pictureBox6.Visible = true;
            formOne.label16.Visible = true;
            formOne.listBox1.Visible = true;

            formOne.pictureBox5.Visible = true;

            formOne.SAVE_ods.Visible = false;
            formOne.button4.Visible = false;
            formOne.label34.Visible = false;
        }


        private void designOdsDetail(Form1 formOne)
        {
            formOne.MaximumSize = new System.Drawing.Size(920, 720);
            formOne.MinimumSize = new System.Drawing.Size(920, 720);

            formOne.label3.Text = "Dettaglio dell'ordine di servizio";
            formOne.label3.ForeColor = System.Drawing.Color.MediumPurple;

            // Rimuovi il pannello del filtro ods
            formOne.Controls.Remove(formOne.panel3);
            // Rimuovi il pannello con la gridView1 degli ods
            formOne.Controls.Remove(formOne.panel4);
            // Rimuovi il pannello del dettaglio dati gev
            formOne.Controls.Remove(formOne.panel8);

            formOne.Controls.Add(formOne.panel7);
        }

        public bool validate(Form1 formOne){

            if (formOne.textBox_odsNumero.Text.Equals("")){
                MessageBox.Show("NUMERO mancante!", "ERRORE");
                return false;
            }

            double dblNumber;
            try {
                dblNumber = Convert.ToDouble(formOne.textBox_odsNumero.Text.Replace(".",","));
            }
            catch (System.FormatException) {
                MessageBox.Show("NUMERO non corretto!", "ERRORE"); 
                return false; 
            }

            //if (dblNumber < 0.0 || dblNumber > 9999.9)
            //{
            //    MessageBox.Show("Il NUMERO deve essere compreso \n tra 0.0 e 9999.9", "ERRORE");
            //    return false; 
            //}

            if (formOne.label31.Text.Equals("")){
                MessageBox.Show("Selezionare un SERVIZIO! ", "ERRORE");
                return false; 
            }

            try { 
                DateTime miaData = Convert.ToDateTime(formOne.textBox_odsDurata.Text);
            } catch (FormatException ) {
                MessageBox.Show("Formato DURATA hh:mm \n (max 23:59)  ", "ERRORE");
                return false;
            }

            return true;
        }


    }
}
