using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace WindowsFormsApplication1
{
    class GeneralDesigner
    {

        public void initializeGridPartecipanti(Form1 formOne) {
            DataGridViewTextBoxColumn gevIdDataGridViewColumn = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn odsIdDataGridViewColumn = new DataGridViewTextBoxColumn();
            DataGridViewLinkColumn cognomeDataGridViewColumn = new DataGridViewLinkColumn();
            DataGridViewTextBoxColumn nomeDataGridViewColumn = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn radioDataGridViewColumn = new DataGridViewTextBoxColumn();
            DataGridViewImageColumn deleteImageDataGridViewColumn = new DataGridViewImageColumn();

            gevIdDataGridViewColumn.DataPropertyName = "gev_ID";
            gevIdDataGridViewColumn.HeaderText = "Gev_ID";
            gevIdDataGridViewColumn.Name = "gevid";
            gevIdDataGridViewColumn.ReadOnly = true;
            gevIdDataGridViewColumn.Visible = false;

            odsIdDataGridViewColumn.DataPropertyName = "ods_ID";
            odsIdDataGridViewColumn.HeaderText = "Ods_ID";
            odsIdDataGridViewColumn.Name = "odsid";
            odsIdDataGridViewColumn.ReadOnly = true;
            odsIdDataGridViewColumn.Visible = false;

            cognomeDataGridViewColumn.DataPropertyName = "Cognome";
            cognomeDataGridViewColumn.HeaderText = "Cognome Gev";
            cognomeDataGridViewColumn.ReadOnly = true;
            cognomeDataGridViewColumn.Width = 200;
            cognomeDataGridViewColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            cognomeDataGridViewColumn.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 9F);

            nomeDataGridViewColumn.DataPropertyName = "Nome";
            nomeDataGridViewColumn.HeaderText = "Nome";
            nomeDataGridViewColumn.Width = 170;
            nomeDataGridViewColumn.ReadOnly = true;

            radioDataGridViewColumn.DataPropertyName = "Numero radio";
            radioDataGridViewColumn.HeaderText = "N. radio";
            radioDataGridViewColumn.Width = 70;
            radioDataGridViewColumn.ReadOnly = true;
            radioDataGridViewColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            deleteImageDataGridViewColumn.Width = 35;
            Image deleteImage = Image.FromFile("risorse\\Delete-icon12.png");
            deleteImageDataGridViewColumn.Image = deleteImage;

            formOne.dataGridView3.Columns.Clear();
            formOne.dataGridView3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
                        gevIdDataGridViewColumn,
                        odsIdDataGridViewColumn,
                        cognomeDataGridViewColumn,
                        nomeDataGridViewColumn,
                        radioDataGridViewColumn,
                        deleteImageDataGridViewColumn
                });
        }

        public void initializeGridListOds(Form1 formOne) {
            DataGridViewTextBoxColumn ods_Id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewLinkColumn numeroDataGridViewColumn = new System.Windows.Forms.DataGridViewLinkColumn();
            DataGridViewTextBoxColumn dataDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn durataDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn tipoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn luogoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn partecipantiDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();

            ods_Id.DataPropertyName = "Ods_id";
            ods_Id.HeaderText = "Ods_ID";
            ods_Id.Name = "odsid";
            ods_Id.ReadOnly = true;
            ods_Id.Visible = false;

            numeroDataGridViewColumn.DataPropertyName = "Numero";
            numeroDataGridViewColumn.HeaderText = "Numero";
            numeroDataGridViewColumn.ReadOnly = true;
            numeroDataGridViewColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            numeroDataGridViewColumn.Width = 80;

            dataDataGridViewTextBoxColumn.DataPropertyName = "Data";
            dataDataGridViewTextBoxColumn.HeaderText = "Data";
            dataDataGridViewTextBoxColumn.ReadOnly = true;
            dataDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            durataDataGridViewTextBoxColumn.DataPropertyName = "Durata";
            durataDataGridViewTextBoxColumn.HeaderText = "Durata";
            durataDataGridViewTextBoxColumn.ReadOnly = true;
            durataDataGridViewTextBoxColumn.DefaultCellStyle.Format = "t";
            durataDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            durataDataGridViewTextBoxColumn.Width = 75;

            tipoDataGridViewTextBoxColumn.DataPropertyName = "tipoServizio";
            tipoDataGridViewTextBoxColumn.HeaderText = "Tipo";
            tipoDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tipoDataGridViewTextBoxColumn.ReadOnly = true;
            tipoDataGridViewTextBoxColumn.Width = 75;

            luogoDataGridViewTextBoxColumn.DataPropertyName = "luogo";
            luogoDataGridViewTextBoxColumn.HeaderText = "Luogo";
            luogoDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            luogoDataGridViewTextBoxColumn.ReadOnly = true;
            luogoDataGridViewTextBoxColumn.Width = 75;

            partecipantiDataGridViewTextBoxColumn.DataPropertyName = "partecipanti";
            partecipantiDataGridViewTextBoxColumn.HeaderText = "Partecipanti";
            partecipantiDataGridViewTextBoxColumn.ReadOnly = true;
            partecipantiDataGridViewTextBoxColumn.Width = 250;

            formOne.dataGridView1.Columns.Clear();
            formOne.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
                                    ods_Id,
                                    numeroDataGridViewColumn,
                                    dataDataGridViewTextBoxColumn,
                                    durataDataGridViewTextBoxColumn,
                                    tipoDataGridViewTextBoxColumn,
                                    luogoDataGridViewTextBoxColumn,
                                    partecipantiDataGridViewTextBoxColumn
                    });
        }

        public void initializeGridOdsEffettuati(Form1 formOne)
        {
            DataGridViewTextBoxColumn ods_Id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewLinkColumn numeroDataGridViewColumn = new System.Windows.Forms.DataGridViewLinkColumn();
            DataGridViewTextBoxColumn dataDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn durataDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn tipoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn descrizioneDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();

            ods_Id.DataPropertyName = "Ods_id";
            ods_Id.HeaderText = "Ods_ID";
            ods_Id.Name = "odsid";
            ods_Id.ReadOnly = true;
            ods_Id.Visible = false;

            numeroDataGridViewColumn.DataPropertyName = "Numero";
            numeroDataGridViewColumn.HeaderText = "Numero";
            numeroDataGridViewColumn.ReadOnly = true;
            numeroDataGridViewColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;

            dataDataGridViewTextBoxColumn.DataPropertyName = "Data";
            dataDataGridViewTextBoxColumn.HeaderText = "Data";
            dataDataGridViewTextBoxColumn.ReadOnly = true;

            durataDataGridViewTextBoxColumn.DataPropertyName = "Durata";
            durataDataGridViewTextBoxColumn.HeaderText = "Durata";
            durataDataGridViewTextBoxColumn.ReadOnly = true;
            durataDataGridViewTextBoxColumn.DefaultCellStyle.Format = "t";
            durataDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            tipoDataGridViewTextBoxColumn.DataPropertyName = "tipoServizio";
            tipoDataGridViewTextBoxColumn.HeaderText = "Tipo";
            tipoDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            tipoDataGridViewTextBoxColumn.ReadOnly = true;

            descrizioneDataGridViewTextBoxColumn.DataPropertyName = "DESCRIZIONE";
            descrizioneDataGridViewTextBoxColumn.HeaderText = "Descrizione";
            descrizioneDataGridViewTextBoxColumn.Width = 250;
            descrizioneDataGridViewTextBoxColumn.ReadOnly = true;

            formOne.dataGridView4.Columns.Clear();
            formOne.dataGridView4.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
                                    ods_Id,
                                    numeroDataGridViewColumn,
                                    dataDataGridViewTextBoxColumn,
                                    durataDataGridViewTextBoxColumn,
                                    tipoDataGridViewTextBoxColumn,
                                    descrizioneDataGridViewTextBoxColumn});

        }

        public void initializeGridListGev(Form1 formOne)
        {
            DataGridViewTextBoxColumn gev_Id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn radioDataGridViewColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewLinkColumn cognomeDataGridViewColumn = new System.Windows.Forms.DataGridViewLinkColumn();
            DataGridViewTextBoxColumn nomeDataGridViewColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn emailDataGridViewColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn telefonoDataGridViewColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn ruoloDataGridViewColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();

            gev_Id.DataPropertyName = "Gev_ID";
            gev_Id.HeaderText = "Gev_ID";
            gev_Id.Name = "gevid";
            gev_Id.ReadOnly = true;
            gev_Id.Visible = false;

            radioDataGridViewColumn.DataPropertyName = "Radio";
            radioDataGridViewColumn.HeaderText = "Radio";
            radioDataGridViewColumn.ReadOnly = true;

            cognomeDataGridViewColumn.DataPropertyName = "Cognome";
            cognomeDataGridViewColumn.HeaderText = "Cognome";
            cognomeDataGridViewColumn.ReadOnly = true;
            cognomeDataGridViewColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;

            nomeDataGridViewColumn.DataPropertyName = "Nome";
            nomeDataGridViewColumn.HeaderText = "Nome";
            nomeDataGridViewColumn.ReadOnly = true;

            emailDataGridViewColumn.DataPropertyName = "email";
            emailDataGridViewColumn.HeaderText = "E-mail";
            emailDataGridViewColumn.Width = 200;
            emailDataGridViewColumn.ReadOnly = true;

            telefonoDataGridViewColumn.DataPropertyName = "fisso";
            telefonoDataGridViewColumn.HeaderText = "Recapito";
            telefonoDataGridViewColumn.ReadOnly = true;

            ruoloDataGridViewColumn.DataPropertyName = "Ruolo";
            ruoloDataGridViewColumn.HeaderText = "Ruolo";
            ruoloDataGridViewColumn.Width = 70;
            ruoloDataGridViewColumn.ReadOnly = true;

            formOne.dataGridView2.Columns.Clear();
            formOne.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
                                    gev_Id,
                                    radioDataGridViewColumn,
                                    cognomeDataGridViewColumn,
                                    nomeDataGridViewColumn,
                                    emailDataGridViewColumn, 
                                    telefonoDataGridViewColumn,
                                    ruoloDataGridViewColumn});
        }

    }
}
