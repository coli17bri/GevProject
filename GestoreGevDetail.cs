using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    class GestoreGevDetail
    {

        public DataRow getGevDetailFromDB(int odsKey)
        {
            String queryString;
            DataSet dsResult = new DataSet();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT Gev_ID, Cognome, Nome, Ruolo, Radio, [Telefono Fisso] as fisso, [Telefono Mobile] as cellulare, [e-mail] as email, Note ";
            queryString = queryString + "FROM tab_DatiAnagrafici ";
            queryString = queryString + "WHERE tab_DatiAnagrafici.Gev_ID=" + odsKey;

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dataAdapter.Fill(dsResult, "DetailGEV");

            connection.Close();
            connection.Dispose();
            return dsResult.Tables["DetailGEV"].Rows[0];
        }

        public DataTable getOdsEffettuati (Form1 formOne, int gevkey) { 
            String queryString;
            DataTable dTable = new DataTable();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT tab_rel_Gev_Ods.Ods_ID,  tab_OrdiniDiServizio.Numero, tab_OrdiniDiServizio.Data, ";
            queryString = queryString + "tab_OrdiniDiServizio.DURATA, tab_OrdiniDiServizio.[Codice servizio] as tipoServizio, tabd_Servizi.DESCRIZIONE ";
            queryString = queryString + "FROM (tab_rel_Gev_Ods INNER JOIN tab_OrdiniDiServizio ON tab_rel_Gev_Ods.Ods_ID = tab_OrdiniDiServizio.ODS_ID) ";
            queryString = queryString + "INNER JOIN tabd_Servizi ON tab_OrdiniDiServizio.[CODICE SERVIZIO] = tabd_Servizi.CODICE ";
            queryString = queryString + "WHERE  tab_rel_Gev_Ods.Gev_ID = " + gevkey + " ";
            queryString = queryString + "AND tab_rel_Gev_Ods.D_END = cDate('31/12/9999') ";
            queryString = queryString + "AND  Data BETWEEN cDate('" + formOne.dateTimePicker1.Value.ToShortDateString() + "') AND cDate('" + formOne.dateTimePicker2.Value.ToShortDateString() + "') ";
            if (formOne.comboBox4.SelectedIndex > 0)
                queryString = queryString + "AND [Codice servizio] = '" + formOne.comboBox4.SelectedItem.ToString() + "' ";
            if (!formOne.textBox11.Text.Equals(""))
                queryString = queryString + "AND Numero = '" + formOne.textBox11.Text + "' ";
            queryString = queryString + "ORDER BY tab_OrdiniDiServizio.Data DESC ";
         
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dTable.Clear();
            dataAdapter.Fill(dTable);

            connection.Close(); 
            connection.Dispose();
            return dTable;
        }

        public String getTotaleOre(Form1 formOne, int gevkey) {
            String strSelectSum;
            double sommaOre = 0.0;

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            strSelectSum = "SELECT sum(tab_OrdiniDiServizio.DURATA) AS [Totale ore] ";
            strSelectSum = strSelectSum + "FROM tab_OrdiniDiServizio INNER JOIN tab_rel_Gev_Ods ON tab_OrdiniDiServizio.ODS_ID = tab_rel_Gev_Ods.Ods_ID ";
            strSelectSum = strSelectSum + "WHERE tab_rel_Gev_Ods.Gev_ID = " + gevkey + " ";
            strSelectSum = strSelectSum + "AND tab_rel_Gev_Ods.D_END = cDate('31/12/9999') ";
            strSelectSum = strSelectSum + "AND  Data BETWEEN cDate('" + formOne.dateTimePicker1.Value.ToShortDateString() + "') AND cDate('" + formOne.dateTimePicker2.Value.ToShortDateString() + "') ";
            if (formOne.comboBox4.SelectedIndex > 0)
                strSelectSum = strSelectSum + "AND [Codice servizio] = '" + formOne.comboBox4.SelectedItem.ToString() + "' ";
            if (!formOne.textBox11.Text.Equals(""))
                strSelectSum = strSelectSum + "AND Numero = '" + formOne.textBox11.Text + "' ";

            OleDbCommand dbCommand = new OleDbCommand(strSelectSum, connection);
            OleDbDataReader dataReader = dbCommand.ExecuteReader();
            if (dataReader.HasRows) {
              dataReader.Read();
              if (!dataReader.IsDBNull(0))
                sommaOre = Convert.ToDouble(dataReader["Totale ore"]);
            } 
            dataReader.Close();

            connection.Close();
            connection.Dispose();

            return GeneralHelper.ConvertDaysToHours(sommaOre);
        }

        public DataSet getComboxList()
        {
            String queryString;
            DataSet dsResult = new DataSet();

            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }
            queryString = "SELECT Codice, Descrizione ";
            queryString = queryString + "FROM tabd_Ruoli ";
            //queryString = queryString + "WHERE Descrizione <> ''";

            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(queryString, connection);
            dataAdapter.Fill(dsResult, "ListaRuoli");

            connection.Close();
            connection.Dispose();
            return dsResult;
        }

        public int updateGevDetail(Form1 formOne, int gevkey)
        {
            String strSQL;
            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            strSQL = "UPDATE tab_DatiAnagrafici ";
            //strSQL = strSQL + "SET Cognome = '" + formOne.textBox_gevCognome.Text + "', "
            //                + "Nome = '" + formOne.textBox_gevNome.Text + "', "
            //                + "Ruolo = '" + formOne.comboBox2.SelectedValue + "', "
            //                + "Radio = '" + formOne.textBox_gevRadio.Text + "', "
            //                + "[Telefono Fisso] = '" + formOne.textBox_gevTelFisso.Text + "', "
            //                + "[Telefono Mobile] = '" + formOne.textBox_gevTelMobile.Text + "', "
            //                + "[e-mail] = '" + formOne.textBox_gevEmail.Text + "', "
            //                + "[Note] = '" + formOne.richTextBox_gevNote.Text + "' ";
            strSQL = strSQL + "SET Cognome = ?, "
                            + "Nome = ?, "
                            + "Ruolo = '" + formOne.comboBox2.SelectedValue + "', "
                            + "Radio = '" + formOne.textBox_gevRadio.Text + "', "
                            + "[Telefono Fisso] = '" + formOne.textBox_gevTelFisso.Text + "', "
                            + "[Telefono Mobile] = '" + formOne.textBox_gevTelMobile.Text + "', "
                            + "[e-mail] = '" + formOne.textBox_gevEmail.Text + "', "
                            + "[Note] = ? ";
            strSQL = strSQL + "WHERE Gev_ID = " + gevkey;

            OleDbCommand dbCommand = new OleDbCommand(strSQL, connection);
            dbCommand.Parameters.Clear();
            dbCommand.Parameters.Add("@pCognome", OleDbType.VarChar, 100).Value = formOne.textBox_gevCognome.Text;
            dbCommand.Parameters.Add("@pNome", OleDbType.VarChar, 100).Value = formOne.textBox_gevNome.Text;
            dbCommand.Parameters.Add("@pNote", OleDbType.VarChar, 300).Value = formOne.richTextBox_gevNote.Text;
            int rowsAffected = dbCommand.ExecuteNonQuery();

            connection.Close();
            connection.Dispose();

            return rowsAffected;
        }

        public void fillCombobox(Form1 formOne)
        {
            formOne.comboBox2.DataSource = this.getComboxList().Tables["ListaRuoli"];
            formOne.comboBox2.DisplayMember = "Descrizione";
            formOne.comboBox2.ValueMember = "Codice";
            GeneralHelper.createMulticolumnCombobox(formOne.comboBox2, 0.3);
        }

        public void fillGevDetailForm(Form1 formOne, int gevIdSelected)
        {
            this.fillCombobox(formOne);
            this.designGevDetail(formOne);

            DataRow datiBaseGev = this.getGevDetailFromDB(gevIdSelected);
            formOne.label_gevId.Text = datiBaseGev.ItemArray.GetValue(0).ToString();
            formOne.textBox_gevCognome.Text = datiBaseGev["Cognome"].ToString();
            formOne.textBox_gevNome.Text = datiBaseGev["Nome"].ToString();
            formOne.comboBox2.SelectedValue = datiBaseGev["Ruolo"].ToString();
            formOne.textBox_gevRadio.Text = datiBaseGev["Radio"].ToString();
            formOne.textBox_gevTelFisso.Text = datiBaseGev["fisso"].ToString();
            formOne.textBox_gevTelMobile.Text = datiBaseGev["cellulare"].ToString();
            formOne.textBox_gevEmail.Text = datiBaseGev["email"].ToString();
            formOne.richTextBox_gevNote.Text = datiBaseGev["Note"].ToString();

            formOne.dataGridView4.DataSource = this.getOdsEffettuati(formOne, gevIdSelected);
            formOne.dataGridView4.ClearSelection();

            formOne.label18.Text = this.getTotaleOre(formOne, gevIdSelected);

            formOne.button6.Visible = false;
            formOne.button5.Visible = false;
            formOne.label8.Visible = false;
        }
        
        public void designGevDetail(Form1 formOne)
        {
            formOne.MaximumSize = new System.Drawing.Size(920, 800);
            formOne.MinimumSize = new System.Drawing.Size(920, 800);

            formOne.label3.Text = "Dati informativi della Gev";
            formOne.label3.ForeColor = System.Drawing.Color.MediumSeaGreen;

            formOne.dateTimePicker1.Value = new DateTime(DateTime.Today.Year, 1, 1);
            formOne.dateTimePicker2.Value = new DateTime(DateTime.Today.Year, 12, 31);
            formOne.comboBox4.SelectedIndex = 0;
            formOne.textBox11.Text = "";

            // Rimuovi il pannello con la gridView2 dell'anagrafica
            formOne.Controls.Remove(formOne.panel6);
            // Rimuovi il pannello del dettaglio ods
            formOne.Controls.Remove(formOne.panel7);

            formOne.Controls.Add(formOne.panel8);
        
        }

    }
}


 //public void ProcessRequest (HttpContext context) {
 //       context.Response.ContentType = "image/jpeg";
 
 //       int id = Convert.ToInt32(context.Request.QueryString["ImageID"]);
 
 //       string cmdText = "SELECT ImageData FROM ImageSave WHERE ImageID = " + id;
 //       string myConnection = "Data Source=Local;Initial Catalog=master;Integrated Security=True";
 //       SqlConnection connection = new SqlConnection(myConnection);
 //       SqlCommand command = new SqlCommand(cmdText, connection);
 //       connection.Open();
 
 //       SqlDataReader reader = command.ExecuteReader();
 //       reader.Read();
 //       byte[] image = (byte[])reader[0];
 //       context.Response.BinaryWrite(image);
 //       reader.Close();
 //       connection.Close();
       
 //   }




