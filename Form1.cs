using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public static String VERSION = "V1.2014";
        public static String START_NUM = "0000.0";
        public static String END_NUM = "9999.9";
        public static int ANNO_ZERO = 2014;

        ReportHelper reportHelper = new ReportHelper();
        GeneralDesigner generalDesigner = new GeneralDesigner();
        GestoreOdsList gestoreOdsList = new GestoreOdsList();
        GestoreGevList gestoreGevList = new GestoreGevList();
        GestoreOdsDetail gestoreOdsDetail = new GestoreOdsDetail();
        GestoreGevDetail gestoreGevDetail = new GestoreGevDetail();
        FormSelezionePartecipanti frmSelPart = new FormSelezionePartecipanti();
        FiltroOds filtroOds = new FiltroOds();

        private bool odsInsertMode = false;
        public void setOdsInsertMode(bool pOdsInsertMode) { odsInsertMode = pOdsInsertMode;}
        public bool getOdsInsertMode() { return odsInsertMode; }
        private bool odsUpdateMode = false;
        public void setOdsUpdateMode(bool pOdsUpdateMode) { odsUpdateMode = pOdsUpdateMode; }
        public bool getOdsUpdateMode() { return odsUpdateMode; }
        private int annoInCorso;

        public Form1()
        {
            InitializeComponent();
            generalDesigner.initializeGridPartecipanti(this);
            generalDesigner.initializeGridListOds(this);
            generalDesigner.initializeGridOdsEffettuati(this);
            generalDesigner.initializeGridListGev(this);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Inizializzazione del filtro
            //textBox1.Text = START_NUM;
            //textBox2.Text = END_NUM;
            annoInCorso = DateTime.Today.Year;
            dateTimePicker4.Value = new DateTime(annoInCorso, 12, 31);
            //this.dateTimePicker4.MaxDate = new System.DateTime(ANNO_RIF, 12, 31, 0, 0, 0, 0);
            //this.dateTimePicker4.MinDate = new System.DateTime(ANNO_RIF-1, 12, 1, 0, 0, 0, 0);
            dateTimePicker3.Value = new DateTime(annoInCorso, DateTime.Today.Month, 1);
            //this.dateTimePicker3.MaxDate = new System.DateTime(ANNO_RIF, 12, 31, 0, 0, 0, 0);
            //this.dateTimePicker3.MinDate = new System.DateTime(ANNO_RIF-1, 12, 1, 0, 0, 0, 0);
            //this.dateTimePicker_odsEffettuato.MaxDate = new System.DateTime(ANNO_RIF, 12, 31, 0, 0, 0, 0);
            //this.dateTimePicker_odsEffettuato.MinDate = new System.DateTime(ANNO_RIF-1, 12, 1, 0, 0, 0, 0);
            //this.dateTimePicker1.MaxDate = new System.DateTime(ANNO_RIF, 12, 31, 0, 0, 0, 0);
            //this.dateTimePicker1.MinDate = new System.DateTime(ANNO_RIF-1, 12, 1, 0, 0, 0, 0);
            //this.dateTimePicker2.MaxDate = new System.DateTime(ANNO_RIF, 12, 31, 0, 0, 0, 0);
            //this.dateTimePicker2.MinDate = new System.DateTime(ANNO_RIF-1, 12, 1, 0, 0, 0, 0);
            AnnoReportComboBox.SelectedIndex = DateTime.Today.Year - ANNO_ZERO;
            gestoreOdsList.fillComboFilter(this);
            gestoreGevList.getRuoliList(this);
            textBox3.Text = "";

            label4_Click(sender, e);
        }

        //////----------- ODS LIST -------------------------------------------------
        // Selezione della voce di menù per la lista degli ods 
        private void label4_Click(object sender, EventArgs e) {
            odsInsertMode = false;
            odsUpdateMode = true;
            refreshOdsTable();
            gestoreOdsList.designOdsList(this);
        }

        private void pictureBox2_Click(object sender, EventArgs e) {
            refreshOdsTable();
        }

        private void refreshOdsTable() {
            this.manageOdsFilterFromList();
            this.dataGridView1.DataSource = gestoreOdsList.getOdsListFromDB(this.filtroOds);
            this.dataGridView1.ClearSelection();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e) {
            manageOdsListClick(e, dataGridView1);
        }

        //////----------- ODS DETAIL -------------------------------------------------
        // Selezione di uno dei partecipanti 
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            manageGevListClick(e, dataGridView3);
        }

        ///////-------------GEV LIST -------------------------------------------
        // Selezione della voce di menù per l'anagrafica 
        private void label7_Click(object sender, EventArgs e) {
            // Lettura dei dati
            refreshGevTable();
            gestoreGevList.designGevList(this);
        }

        private void refreshGevTable() {
            this.dataGridView2.DataSource = gestoreGevList.getGevListFromDB(this);
            this.dataGridView2.ClearSelection();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            manageGevListClick(e, dataGridView2);
        }

        //////----------- GEV DETAIL -------------------------------------------------
        // Selezione di uno degli ods effettuati 
        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            manageOdsListClick(e, dataGridView4);
        }

        private void pictureBox_gevFoto_Click(object sender, EventArgs e)
        {
            if (Control.ModifierKeys == Keys.Alt)
            {
                MessageBox.Show("Alt + click!");
            }
            else
                MessageBox.Show("Other click!");


        }

        ///////// ------------------COMMON---------------------------------------------------

        private void manageGevListClick(DataGridViewCellEventArgs cellEvent, DataGridView dataGridView)
        {
            // is NOT a header cell
            if (cellEvent.RowIndex != -1)
            {
                int gevIdSelected = gestoreGevList.retrieveGevID(cellEvent, dataGridView);
                if (gevIdSelected != -1)
                {
                    if (isALinkCell(cellEvent, dataGridView))
                    {
                        gestoreGevDetail.fillGevDetailForm(this, gevIdSelected);
                    }
                    else if (isAnImageCell(cellEvent, dataGridView))
                    {
                        int rowsAffected = gestoreOdsDetail.deleteOdsPartecipante(gevIdSelected, Convert.ToInt32(this.label_odsId.Text));
                        if (rowsAffected != 0)
                        {
                            this.dataGridView3.DataSource = gestoreOdsDetail.getOdsPartecipanti(Convert.ToInt32(this.label_odsId.Text));
                            this.dataGridView3.ClearSelection();
                        }
                    }
                    else if (isAButtonCell(cellEvent, dataGridView))
                    { }
                }
            }
        }

        private void manageOdsListClick(DataGridViewCellEventArgs cellEvent, DataGridView dataGridView)
        {
            if (cellEvent.RowIndex != -1)
            {
                int odsIdSelected = gestoreOdsList.retrieveOdsID(cellEvent, dataGridView);
                if (odsIdSelected != -1)
                {
                    if (isALinkCell(cellEvent, dataGridView))
                    {
                        gestoreOdsDetail.fillOdsDetailForm(this, odsIdSelected);
                    }
                    else if (isAButtonCell(cellEvent, dataGridView))
                    {  }
                }
            }
        }
 
        private bool isALinkCell(DataGridViewCellEventArgs cellEvent, DataGridView dataGridView)
        {
            if (dataGridView.Columns[cellEvent.ColumnIndex] is DataGridViewLinkColumn )
            { return true; }
            else { return false; }
        }

        private bool isAButtonCell(DataGridViewCellEventArgs cellEvent, DataGridView dataGridView)
        {
            if (dataGridView.Columns[cellEvent.ColumnIndex] is DataGridViewButtonColumn )
            { return true; }
            else { return (false); }
        }

        private bool isAnImageCell(DataGridViewCellEventArgs cellEvent, DataGridView dataGridView)
        {
            if (dataGridView.Columns[cellEvent.ColumnIndex] is DataGridViewImageColumn )
            { return true; }
            else { return (false); }
        }

        private void manageOdsFilterFromList() {
            int comp = dateTimePicker3.Value.CompareTo(dateTimePicker4.Value);
            if (comp > 0) {
                DateTime tempDate = dateTimePicker3.Value;
                dateTimePicker3.Value = dateTimePicker4.Value;
                dateTimePicker4.Value = tempDate;
            }
            filtroOds.setDataFrom(dateTimePicker3.Value);
            filtroOds.setDataTo(dateTimePicker4.Value);
            // ---
            bool numFromOk = true;
            bool numToOK = true;
            String numFrom = textBox1.Text.Trim();
            String numTo = textBox2.Text.Trim();

            if (numFrom.Equals("")) filtroOds.setNumeroFrom(START_NUM);
            if (numTo.Equals("")) filtroOds.setNumeroTo(END_NUM);

            if (!numFrom.Equals(""))
            {
                try { Convert.ToDouble(numFrom); }
                catch (System.FormatException) { MessageBox.Show("Valore numerico non corretto", "ERRORE"); numFromOk = false; }
                if (numFromOk) {
                    numFrom = GeneralHelper.formatNumOds(numFrom);
                    filtroOds.setNumeroFrom(numFrom);
                }
            }

            if (!numTo.Equals(""))
            {
                try { Convert.ToDouble(numTo); }
                catch (System.FormatException) { MessageBox.Show("Valore numerico non corretto", "ERRORE"); numToOK = false; }
                if (numToOK) {
                    numTo = GeneralHelper.formatNumOds(numTo);
                    filtroOds.setNumeroTo(numTo);
                }
            }
            //--
            filtroOds.setTipo("-1");
            if (this.comboBox1.SelectedIndex > 0)
              filtroOds.setTipo(this.comboBox1.SelectedItem.ToString());
            filtroOds.setLuogo("-1");
            if (this.comboBox3.SelectedIndex > 0)
                filtroOds.setLuogo(this.comboBox3.SelectedItem.ToString());
        }


        private void comboBox_odsTipo_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.label31.Text = "";
            if (this.comboBox_odsTipo.SelectedValue != null)
                this.label31.Text = this.comboBox_odsTipo.SelectedValue.ToString();
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }

        private void comboBox_odsLocalita_SelectedIndexChanged(object sender, EventArgs e) {
            this.label32.Text = "";
            if (this.comboBox_odsLocalita.SelectedValue != null )
                this.label32.Text = this.comboBox_odsLocalita.SelectedValue.ToString();
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            frmSelPart.setOdsId(this.label_odsId.Text);
            if (frmSelPart.ShowDialog() == DialogResult.OK)
            {
                this.dataGridView3.DataSource = gestoreOdsDetail.getOdsPartecipanti(Convert.ToInt32(this.label_odsId.Text));
                this.dataGridView3.ClearSelection();
            }
        }

        // DELETE the ods
        private void pictureBox5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Desideri CANCELLARE questo ods?", "ODS Application",
                      MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                gestoreOdsDetail.deleteOdsDetail(Convert.ToInt32(this.label_odsId.Text));
                // e... torna alla lista
                label4_Click(sender, e);
            }

        }

        // CREATE a new ods
        private void initNewOds()
        {
            odsInsertMode = true;
            odsUpdateMode = false;

            // Impostare i combobox
            gestoreOdsDetail.fillCombobox(this);

            // Inizializzare i campi 

            this.label_odsId.Text = "";
            this.textBox_odsNumero.Text = "";
            this.dateTimePicker_odsEffettuato.Value = DateTime.Today;
            this.textBox_odsDurata.Text = "04:00";
            this.comboBox_odsTipo.SelectedValue = "";
            this.comboBox_odsLocalita.SelectedValue = "";
            this.richTextBox1.Text = "";
            this.checkBox2.Checked = false;
            this.checkBox1.Checked = false;

            this.dataGridView3.Visible = false;
            this.label15.Visible = false;
            this.pictureBox6.Visible = false;
            this.label16.Visible = false;
            this.listBox1.Visible = false;

            this.pictureBox5.Visible = false;
        }

        private void SAVE_ods_Click(object sender, EventArgs e)
        {
            if (gestoreOdsDetail.validate(this)){
                if (this.odsInsertMode) {

                    gestoreOdsDetail.insertOdsDetail(this);
                    this.odsInsertMode = false;
                    this.odsUpdateMode = true;

                    this.dataGridView3.Visible = true;
                    DataTable dTable = (DataTable)this.dataGridView3.DataSource;
                    if (dTable != null) dTable.Clear();
                    this.label15.Visible = true;
                    this.pictureBox6.Visible = true;
                    this.label16.Visible = true;
                    this.listBox1.Visible = true;

                    this.pictureBox5.Visible = true;
                }
                else if (this.odsUpdateMode) {
                    gestoreOdsDetail.updateOdsDetail(this, Convert.ToInt32(this.label_odsId.Text));
                }
                this.SAVE_ods.Visible = false;
                this.button4.Visible = false;
                this.label34.Visible = true;
            }
        }

        private void textBox_odsNumero_TextChanged(object sender, EventArgs e) {
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }
        private void textBox_odsDurata_TextChanged(object sender, EventArgs e) {
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }

        private void dateTimePicker_odsEffettuato_ValueChanged(object sender, EventArgs e){
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e) {
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e) {
            this.SAVE_ods.Visible = true;
            this.button4.Visible = true;
            this.label34.Visible = false;
        }

        // Annulla in ods
        private void button4_Click(object sender, EventArgs e)
        {
            label4_Click(sender, e);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox3.Checked)
            {
                dateTimePicker3.Value = new DateTime(DateTime.Today.Year, 1, 1);
                dateTimePicker4.Value = new DateTime(DateTime.Today.Year, 12, 31);
                checkBox3.Checked = false;
            }
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            checkBox3.Checked = true;
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            checkBox3.Checked = true;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox4.Checked) {
                textBox1.Text = "";
                textBox2.Text = "";
                checkBox4.Checked = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            checkBox4.Checked = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            checkBox4.Checked = true;
        }

        // Reset
        private void button3_Click(object sender, EventArgs e)
        {
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            this.pictureBox2_Click(sender, e);
        }

        // NEW ods from list
        private void pictureBox7_Click(object sender, EventArgs e)
        {
            // Rimuovi il pannello del filtro ods
            this.Controls.Remove(this.panel3);
            // Rimuovi il pannello con la gridView1 degli ods
            this.Controls.Remove(this.panel4);

            // abilito il dettaglio ods
            this.Controls.Add(this.panel7);
            // eseguo il tasto di 'create' del dettaglio
            initNewOds();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox5.Checked = this.comboBox1.SelectedIndex > 0 ? true : false;
            if (this.comboBox1.SelectedIndex > 0) pictureBox2_Click(sender, e);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox6.Checked = this.comboBox3.SelectedIndex > 0 ? true : false;
            if (this.comboBox3.SelectedIndex > 0) pictureBox2_Click(sender, e);
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox5.Checked)
            {
                this.comboBox1.SelectedIndex = 0;
            }

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox6.Checked) {
                this.comboBox3.SelectedIndex = 0;
            }
        }

        // Aggiorna gli ods effettuati applicando i filtri
        private void button2_Click(object sender, EventArgs e)
        {
            int gevIdSelected = Convert.ToInt32(this.label_gevId.Text);
            String number = textBox11.Text.Trim();
            bool numberOk = true;

            if (!number.Equals(""))
            {
                try { Convert.ToDouble(number); }
                catch (System.FormatException) { MessageBox.Show("Valore numerico non corretto", "ERRORE"); numberOk = false; }
                if (numberOk)
                    this.textBox11.Text = GeneralHelper.formatNumOds(number);
            }

            if (numberOk)
            {
              this.dataGridView4.DataSource = gestoreGevDetail.getOdsEffettuati(this, gevIdSelected);
              this.dataGridView4.ClearSelection();
              this.label18.Text = gestoreGevDetail.getTotaleOre(this, gevIdSelected);
            }
        }

        // Reset i filtri degli ods effettuati
        private void button1_Click(object sender, EventArgs e)
        {
            this.dateTimePicker1.Value = new DateTime(DateTime.Today.Year, 1, 1);
            this.dateTimePicker2.Value = new DateTime(DateTime.Today.Year, 12, 31);
            this.comboBox4.SelectedIndex = 0;
            this.textBox11.Text = "";
            this.button2_Click(sender, e);
        }

        // Salva dettagli GEV
        private void button6_Click(object sender, EventArgs e)
        {
            gestoreGevDetail.updateGevDetail(this, Convert.ToInt32(this.label_gevId.Text));
            this.button6.Visible = false;
            this.button5.Visible = false;
            this.label8.Visible = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.label7_Click(sender, e);
        }

        // 
        private void textBox_gevCognome_TextChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void textBox_gevNome_TextChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void textBox_gevRadio_TextChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void textBox_gevEmail_TextChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void textBox_gevTelFisso_TextChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void textBox_gevTelMobile_TextChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void richTextBox_gevNote_TextChanged(object sender, EventArgs e)
        {
            this.button6.Visible = true;
            this.button5.Visible = true;
            this.label8.Visible = false;
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            refreshGevTable();
        }
        private void textBox3_TextChanged(object sender, EventArgs e) {
            refreshGevTable();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            refreshGevTable();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshGevTable();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.textBox3.Text = "";
            this.textBox4.Text = "";
            this.textBox5.Text = "";
            this.textBox6.Text = "";
            this.textBox7.Text = "";
            this.comboBox5.SelectedIndex = 0;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
              refreshGevTable();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            refreshGevTable();
        }

        private void RepOrePerMese_Click(object sender, EventArgs e)
        {
            FormAttesa frmAttesa = new FormAttesa();
            frmAttesa.Location = new Point(this.Location.X + 20, this.Location.Y + this.Size.Height / 2);
            frmAttesa.Show();
            reportHelper.ReportOrePerMese(new DateTime(AnnoReportComboBox.SelectedIndex + ANNO_ZERO, 1, 1), new DateTime(AnnoReportComboBox.SelectedIndex + ANNO_ZERO, 12, 31));
            frmAttesa.Close();
        }

        private void RepOrePerServizio_Click(object sender, EventArgs e)
        {
            FormAttesa frmAttesa = new FormAttesa();
            frmAttesa.Location = new Point(this.Location.X + 20, this.Location.Y + this.Size.Height / 2);
            frmAttesa.Show();
            reportHelper.ReportOrePerServizio(new DateTime(AnnoReportComboBox.SelectedIndex + ANNO_ZERO, 1, 1), new DateTime(AnnoReportComboBox.SelectedIndex + ANNO_ZERO, 12, 31));
            frmAttesa.Close();
        }

        private void RepFrequenzaZone_Click(object sender, EventArgs e)
        {
            FormAttesa frmAttesa = new FormAttesa();
            frmAttesa.Location = new Point(this.Location.X + 20, this.Location.Y + this.Size.Height / 2);
            frmAttesa.Show();
            reportHelper.ReportFrequenzaZone(new DateTime(AnnoReportComboBox.SelectedIndex + ANNO_ZERO, 1, 1), new DateTime(AnnoReportComboBox.SelectedIndex + ANNO_ZERO, 12, 31));
            frmAttesa.Close();
        }

        // Search image botton
        private void pictureBox3_Click(object sender, EventArgs e)
        {
            bool isCorrect = true;
            try
            {
              Convert.ToDouble(this.textBox8.Text);
            }
            catch (System.FormatException)
            {
                MessageBox.Show("NUMERO non corretto!", "ERRORE");
                isCorrect = false;
            }

            if (isCorrect)
            {
                int odsIdent = gestoreOdsDetail.searchOds(this.textBox8.Text);
                if (odsIdent != -1)
                {
                    this.Controls.Remove(this.panel6);
                    gestoreOdsDetail.fillOdsDetailForm(this, odsIdent);
                    this.textBox8.Text = "";
                }
                else MessageBox.Show("NUMERO non presente.", "AVVISO");
            }
        }

        private void textBox8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                pictureBox3_Click((object)sender, (EventArgs)e);
            }
        }

        //private void button8_Click(object sender, EventArgs e)
        //{
        //    iTextSharp.text.Document document = new iTextSharp.text.Document();
        //    try
        //    {
        //        PdfWriter.GetInstance(document,
        //            new FileStream(Directory.GetCurrentDirectory() + "\\TestElencoOds.pdf", FileMode.Create));
        //        document.Open();
        //        foreach (DataGridViewRow row in this.dataGridView1.Rows)
        //        {
        //            string rowContent = "";
        //            for (int i = 0; i < row.Cells.Count; i++)
        //            {
        //                rowContent += row.Cells[i].Value + " | ";
        //            }
        //            document.Add(new Phrase(rowContent + "\n"));
        //        }
        //    }
        //    catch (DocumentException de)
        //    {
        //        Console.WriteLine("error" + de.Message);
        //    }
        //    catch (IOException ioe)
        //    {
        //        Console.WriteLine("error " + ioe.Message);
        //    }
        //    document.Close();
        //}

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Boolean typeSelected = this.comboBox1.SelectedIndex > 0;
            reportHelper.excelListaServizi(this.dataGridView1, dateTimePicker3, dateTimePicker4, typeSelected);
        }

        //private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        //{
        //    Bitmap bm = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);
        //    dataGridView1.DrawToBitmap(bm, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));
        //    e.Graphics.DrawImage(bm, 0, 0);
        //}



    }
}
