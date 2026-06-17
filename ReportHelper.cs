using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;

using ExcelRef = Microsoft.Office.Interop.Excel;
using BordersRef = Microsoft.Office.Interop.Excel.XlBordersIndex;
using LineStyleRef = Microsoft.Office.Interop.Excel.XlLineStyle;
using BordersWeightRef = Microsoft.Office.Interop.Excel.XlBorderWeight;
using ColorRef = Microsoft.Office.Interop.Excel.XlRgbColor;

using System.Reflection;
using System.Data.OleDb;



namespace WindowsFormsApplication1
{
    class ReportHelper
    {

        public void ReportFrequenzaZone(DateTime fromDate, DateTime toDate)
        {
            String queryString;
            int shiftRiga, shiftColonna, riga, colonna;
            int totaleMesiAnno = 12;
            int totaleRecordZone = 0;
            Dictionary<String, int> meseColonna = new Dictionary<String, int>();
            Dictionary<String, int> monthTotFreq = new Dictionary<String, int>();
            Dictionary<String, int> zonaRiga = new Dictionary<String, int>();

            ExcelRef.Application xlAppl;
            ExcelRef.Workbook xlWorkbook;
            ExcelRef.Worksheet xlWorksheet;

            xlAppl = new ExcelRef.Application();
            xlWorkbook = xlAppl.Workbooks.Add(Missing.Value);
            xlWorksheet = xlWorkbook.ActiveSheet;
            //xlAppl.Visible = true;
            CultureInfo attuale = Thread.CurrentThread.CurrentUICulture;
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("it-IT");

            // ------------------------------------
            // --------- Scrivere nelle colonne i mesi dell'anno e mettere a 7 la loro larghezza
            shiftColonna = 3;
            riga = 3;
            String meseCode;
            for (int index = 1; index <= totaleMesiAnno; index++)
            {
                meseCode = index.ToString();
                colonna = index + shiftColonna;
                xlWorksheet.Cells[riga, colonna] = CultureInfo.CurrentUICulture.DateTimeFormat.GetAbbreviatedMonthName(index) + ".";
                xlWorksheet.Cells[riga, colonna].Font.Bold = true;
                xlWorksheet.Columns[colonna].ColumnWidth = 7;
                borderDraw(xlWorksheet, riga, colonna);

                meseColonna.Add(meseCode, colonna);
                monthTotFreq.Add(meseCode, 0);
            }
            meseColonna.Add("keyColTot", shiftColonna);

            // ------------------------------------
            // --------- Elencare le zone 
            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT Codice, Descrizione FROM tabd_Zone ";
            queryString = queryString + "ORDER BY Codice ";

            OleDbCommand dbCommand = new OleDbCommand(queryString, connection);
            OleDbDataReader dataReader = dbCommand.ExecuteReader();
            shiftRiga = 4;
            colonna = 1;
            while (dataReader.Read())
            {
                totaleRecordZone++;
                riga = totaleRecordZone + shiftRiga;
                xlWorksheet.Cells[riga, colonna] = "'" + dataReader["Codice"].ToString();
                xlWorksheet.Cells[riga, colonna + 1] = dataReader["Descrizione"].ToString();
                borderDraw(xlWorksheet, riga, colonna);
                borderDraw(xlWorksheet, riga, colonna + 1);

                zonaRiga.Add(dataReader["Codice"].ToString(), riga);
            }
            xlWorksheet.Columns[colonna].ColumnWidth = 6;
            xlWorksheet.Columns[colonna + 1].ColumnWidth = 75;
            xlWorksheet.Columns[colonna + 1].WrapText = true;
            dataReader.Close();

            // Mostrare una griglia e il titolo
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlInsideHorizontal).Weight = BordersWeightRef.xlHairline;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlInsideVertical).Weight = BordersWeightRef.xlHairline;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeTop).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeRight).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeBottom).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeLeft).Weight = BordersWeightRef.xlMedium;

            xlWorksheet.Cells[1, 1].Value = "Prospetto frequenza zone";
            xlWorksheet.Cells[1, 1].Font.Size = 24;
            //xlWorksheet.Cells[2, 1].Value = fromDate.ToLongDateString() ;
            //xlWorksheet.Cells[2, 1].HorizontalAlignment = ExcelRef.Constants.xlCenter;
            //xlWorksheet.Cells[2, 1].Font.Size = 16;
            xlWorksheet.Cells[2, 2].Value = "dal " + fromDate.ToShortDateString() + " al " + toDate.ToShortDateString();
            xlWorksheet.Cells[2, 2].HorizontalAlignment = ExcelRef.Constants.xlCenter;
            xlWorksheet.Cells[2, 2].Font.Size = 16;

            // ------------------------------------
            // Estrazione , elaborazione e visualizzazione dei dati
            queryString = "SELECT tab_OrdiniDiServizio.[CODICE ZONA] AS zonaID, datepart('m',tab_OrdiniDiServizio.[Data]) AS numMese ";
            queryString = queryString + "FROM tab_OrdiniDiServizio ";
            queryString = queryString + "WHERE tab_OrdiniDiServizio.[CODICE ZONA]<>'' ";
            queryString = queryString + "AND tab_OrdiniDiServizio.D_END = cDate('31/12/9999') ";
            queryString = queryString + "ORDER BY tab_OrdiniDiServizio.[CODICE ZONA], tab_OrdiniDiServizio.Data; ";

            dbCommand = new OleDbCommand(queryString, connection);
            dataReader = dbCommand.ExecuteReader();

            String currentZona, currentMonth;
            Boolean eof = !dataReader.Read();
            int freqZona, zonaTotFreq;
            Boolean singolo, lastSingolo;
            while (!eof)
            {
                currentZona = dataReader["zonaID"].ToString();
                currentMonth = dataReader["numMese"].ToString();
                freqZona = 1;
                singolo = true;
                lastSingolo = true;
                zonaTotFreq = 0;
                riga = zonaRiga[currentZona];

                eof = !dataReader.Read();
                if (!eof)
                {
                    while (dataReader["zonaID"].ToString().Equals(currentZona))
                    {
                        singolo = false;
                        while (dataReader["numMese"].ToString().Equals(currentMonth)
                                && dataReader["zonaID"].ToString().Equals(currentZona))
                        {
                            freqZona++;
                            eof = !dataReader.Read();
                            lastSingolo = false;
                            if (eof) break;
                        }
                        // scrivi nel foglio
                        colonna = meseColonna[currentMonth];
                        xlWorksheet.Cells[riga, colonna].Value = freqZona;
                        zonaTotFreq += freqZona;
                        monthTotFreq[currentMonth] += freqZona;

                        // passa al mese successivo
                        if (!eof)
                        {
                            if (dataReader["zonaID"].ToString().Equals(currentZona))
                            {
                                currentMonth = dataReader["numMese"].ToString();
                                freqZona = 1;
                                eof = !dataReader.Read();
                                lastSingolo = true;
                            }
                        }
                        if (eof) break;
                    }
                }

                // caso di un solo ods
                if (singolo || lastSingolo)
                {
                    colonna = meseColonna[currentMonth];
                    xlWorksheet.Cells[riga, colonna].Value = freqZona;
                    zonaTotFreq += freqZona;
                    monthTotFreq[currentMonth] += freqZona;
                }
                colonna = meseColonna["keyColTot"];
                xlWorksheet.Cells[riga, colonna].Value = zonaTotFreq;
                xlWorksheet.Cells[riga, colonna].Font.ColorIndex = 3;
                xlWorksheet.Cells[riga, colonna].HorizontalAlignment = ExcelRef.Constants.xlCenter;            }
            dataReader.Close();

            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, shiftColonna]].Borders(BordersRef.xlEdgeTop).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, shiftColonna]].Borders(BordersRef.xlEdgeBottom).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordZone + shiftRiga, shiftColonna]].Borders(BordersRef.xlInsideHorizontal).LineStyle = LineStyleRef.xlContinuous;

            // ------------------------------------
            // Lettura del dictionary per i totali mensili
            int coco = shiftColonna + 1;
            double totaleAnnuo = 0;
            foreach (KeyValuePair<String, int> kvp in monthTotFreq)
            {
                xlWorksheet.Cells[shiftRiga, coco] = kvp.Value;
                xlWorksheet.Cells[shiftRiga, coco].Font.ColorIndex = 30;
                coco = coco + 1;
                totaleAnnuo = totaleAnnuo + kvp.Value;
            }
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeLeft).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeRight).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlInsideVertical).LineStyle = LineStyleRef.xlContinuous;

            xlWorksheet.Cells[shiftRiga, shiftColonna] = totaleAnnuo;
            xlWorksheet.Cells[shiftRiga, shiftColonna].Font.Color = ColorRef.rgbBlue;
            xlWorksheet.Cells[shiftRiga, shiftColonna].HorizontalAlignment = ExcelRef.Constants.xlCenter;

            xlAppl.Visible = true;

            Thread.CurrentThread.CurrentUICulture = attuale;
            releaseObject(xlWorksheet);
            releaseObject(xlWorkbook);
            releaseObject(xlAppl);
        }


        public void ReportOrePerServizio(DateTime fromDate, DateTime toDate)
        {
            String queryString;
            int shiftRiga, shiftColonna, riga, colonna;
            int totaleRecordServizi = 0;
            Dictionary<String, int> servizioColonna = new Dictionary<String, int>();
            Dictionary<String, double> serviceTotHours = new Dictionary<String, double>();
            Dictionary<int, int> gevRiga = new Dictionary<int, int>();

            ExcelRef.Application xlAppl;
            ExcelRef.Workbook xlWorkbook;
            ExcelRef.Worksheet xlWorksheet;

            xlAppl = new ExcelRef.Application();
            xlWorkbook = xlAppl.Workbooks.Add(Missing.Value);
            xlWorksheet = xlWorkbook.ActiveSheet;
            //xlAppl.Visible = true;

            CultureInfo attuale = Thread.CurrentThread.CurrentUICulture;
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("it-IT");

            // ------------------------------------
            // --------- Scrivere nelle colonne i servizi e mettere a 8 la loro larghezza
            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT Codice, Descrizione FROM tabd_Servizi ";
            queryString = queryString + "ORDER BY Codice ";

            OleDbCommand dbCommand = new OleDbCommand(queryString, connection);
            OleDbDataReader dataReader = dbCommand.ExecuteReader();

            shiftColonna = 3;
            riga = 3;
            String servCode;
            while (dataReader.Read())
            {
                totaleRecordServizi++;
                colonna = totaleRecordServizi + shiftColonna;
                servCode = dataReader["Codice"].ToString();
                xlWorksheet.Cells[riga, colonna] = "'" + servCode;
                //xlWorksheet.Cells[riga, colonna].NumberFormat = "@";
                xlWorksheet.Cells[riga, colonna].Font.Bold = true;
                xlWorksheet.Columns[colonna].ColumnWidth = 8;
                borderDraw(xlWorksheet, riga, colonna);
                xlWorksheet.Cells[riga - 1, colonna] = dataReader["Descrizione"].ToString();
                xlWorksheet.Cells[riga - 1, colonna].RowHeight = 80;
                xlWorksheet.Cells[riga - 1, colonna].Orientation = 90;
                xlWorksheet.Cells[riga - 1, colonna].WrapText = true;

                servizioColonna.Add(servCode, colonna);
                serviceTotHours.Add(servCode, 0.0);
            }
            servizioColonna.Add("keyColTot", shiftColonna);
            dataReader.Close();

            // ------------------------------------
            // --------- Elencare nelle prime 2 colonne i cognomi e nomi
            shiftRiga = 4;
            int totaleRecordGev = ElencaCognomiNomi(xlWorksheet, gevRiga, shiftRiga);

            // ------------------------------------
            // Impostare griglia dati e titolo
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlInsideHorizontal).Weight = BordersWeightRef.xlHairline;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlInsideVertical).Weight = BordersWeightRef.xlHairline;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlEdgeTop).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlEdgeRight).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlEdgeBottom).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlEdgeLeft).Weight = BordersWeightRef.xlMedium;

            xlWorksheet.Cells[1, 1].Value = "Prospetto servizi";
            xlWorksheet.Cells[1, 1].Font.Size = 24;
            xlWorksheet.Cells[2, 1].Value = fromDate;
            xlWorksheet.Cells[2, 1].HorizontalAlignment = ExcelRef.Constants.xlCenter;
            xlWorksheet.Cells[2, 1].Font.Size = 16;
            xlWorksheet.Cells[2, 2].Value = toDate;
            xlWorksheet.Cells[2, 2].HorizontalAlignment = ExcelRef.Constants.xlCenter;
            xlWorksheet.Cells[2, 2].Font.Size = 16;

            // ------------------------------------
            // Elaborazione ed inserimento DATI
            queryString = "SELECT tab_DatiAnagrafici.Gev_id AS GevID, tab_DatiAnagrafici.Cognome, tab_OrdiniDiServizio.[CODICE SERVIZIO] AS ServiceCode, tab_OrdiniDiServizio.DURATA AS Ore ";
            queryString = queryString + "FROM tab_OrdiniDiServizio INNER JOIN (tab_rel_Gev_Ods INNER JOIN (tab_DatiAnagrafici INNER JOIN tabd_Ruoli ON tab_DatiAnagrafici.ruolo = tabd_Ruoli.Codice) ";
            queryString = queryString + "ON tab_rel_Gev_Ods.Gev_ID = tab_DatiAnagrafici.Gev_ID ) ON tab_OrdiniDiServizio.ODS_ID = tab_rel_Gev_Ods.Ods_ID ";
            queryString = queryString + "WHERE tab_OrdiniDiServizio.Data BETWEEN  cDate('" + fromDate.ToShortDateString() + "') AND cDate('" + toDate.ToShortDateString() + "') ";
            queryString = queryString + "AND tab_rel_Gev_Ods.D_END = cDate('31/12/9999') ";
            queryString = queryString + "AND NOT tabd_Ruoli.ToIgnore ";
            queryString = queryString + "ORDER BY tab_DatiAnagrafici.Gev_id, tab_OrdiniDiServizio.[CODICE SERVIZIO]; ";

            dbCommand = new OleDbCommand(queryString, connection);
            dataReader = dbCommand.ExecuteReader();

            int currentGev;
            String currentService;
            double totOre, gevTotHours;
            Boolean singolo, lastSingolo;
            Boolean eof = !dataReader.Read();
            while (!eof)
            {
                currentGev = Convert.ToInt32(dataReader["GevID"]);
                currentService = dataReader["ServiceCode"].ToString();
                totOre = ((DateTime)dataReader["Ore"]).ToOADate();
                singolo = true;
                lastSingolo = true;
                gevTotHours = 0.0;
                riga = gevRiga[currentGev];

                eof = !dataReader.Read();
                if (!eof)
                {
                    while (Convert.ToInt32(dataReader["GevID"]) == currentGev)
                    {
                        singolo = false;
                        while (dataReader["ServiceCode"].ToString().Equals(currentService)
                                && Convert.ToInt32(dataReader["GevID"]) == currentGev)
                        {
                            totOre = totOre + ((DateTime)dataReader["Ore"]).ToOADate();
                            eof = !dataReader.Read();
                            lastSingolo = false;
                            if (eof) break;
                        }
                        // scrivi nel foglio
                        colonna = servizioColonna[currentService];
                        xlWorksheet.Cells[riga, colonna].Value = GeneralHelper.ConvertDaysToHours(totOre);
                        xlWorksheet.Cells[riga, colonna].NumberFormat = "[h]:mm";
                        gevTotHours = gevTotHours + totOre;
                        serviceTotHours[currentService] = serviceTotHours[currentService] + totOre;

                        // passa al servizio successivo
                        if (!eof)
                        {
                            if (Convert.ToInt32(dataReader["GevID"]) == currentGev)
                            {
                                currentService = dataReader["ServiceCode"].ToString();
                                totOre = ((DateTime)dataReader["Ore"]).ToOADate();
                                eof = !dataReader.Read();
                                lastSingolo = true;
                            }
                        }
                        if (eof) break;
                    }
                }

                // caso di un solo ods
                if (singolo || lastSingolo)
                {
                    colonna = servizioColonna[currentService];
                    xlWorksheet.Cells[riga, colonna].Value = GeneralHelper.ConvertDaysToHours(totOre);
                    xlWorksheet.Cells[riga, colonna].NumberFormat = "[h]:mm";
                    gevTotHours = gevTotHours + totOre;
                    serviceTotHours[currentService] = serviceTotHours[currentService] + totOre;
                }
                colonna = servizioColonna["keyColTot"];
                xlWorksheet.Cells[riga, colonna].Value = GeneralHelper.ConvertDaysToHours(gevTotHours);
                xlWorksheet.Cells[riga, colonna].NumberFormat = "[h]:mm";
                xlWorksheet.Cells[riga, colonna].Font.ColorIndex = 3;
            }
            dataReader.Close();

            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, shiftColonna]].Borders(BordersRef.xlEdgeTop).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, shiftColonna]].Borders(BordersRef.xlEdgeBottom).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, shiftColonna]].Borders(BordersRef.xlInsideHorizontal).LineStyle = LineStyleRef.xlContinuous;

            // ------------------------------------
            // visualizzazione dei totali per servizio e totale annuo
            int coco = shiftColonna + 1;
            double totaleAnnuo = 0;
            foreach (KeyValuePair<String, double> kvp in serviceTotHours)
            {
                xlWorksheet.Cells[shiftRiga, coco] = GeneralHelper.ConvertDaysToHours(kvp.Value);
                xlWorksheet.Cells[shiftRiga, coco].NumberFormat = "[h]:mm";
                xlWorksheet.Cells[shiftRiga, coco].Font.ColorIndex = 30;
                coco = coco + 1;
                totaleAnnuo = totaleAnnuo + kvp.Value;
            }
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlEdgeLeft).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlEdgeRight).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleRecordServizi + shiftColonna]].Borders(BordersRef.xlInsideVertical).LineStyle = LineStyleRef.xlContinuous;

            xlWorksheet.Cells[shiftRiga, shiftColonna] = GeneralHelper.ConvertDaysToHours(totaleAnnuo);
            xlWorksheet.Cells[shiftRiga, shiftColonna].Font.Color = ColorRef.rgbBlue;
            xlWorksheet.Cells[shiftRiga, shiftColonna].NumberFormat = "[h]:mm";

            xlAppl.Visible = true;

            Thread.CurrentThread.CurrentUICulture = attuale;
            releaseObject(xlWorksheet);
            releaseObject(xlWorkbook);
            releaseObject(xlAppl);
        }

        public void ReportOrePerMese(DateTime fromDate, DateTime toDate)
        {
            String queryString;
            int totaleRecordGev = 0;

            Dictionary<String, int> meseColonna = new Dictionary<string, int>();
            Dictionary<String, double> monthTotHours = new Dictionary<string, double>();
            int shiftRiga, shiftColonna, riga, colonna;
            int totaleMesiAnno = 12;

            ExcelRef.Application xlAppl;
            ExcelRef.Workbook xlWorkbook;
            ExcelRef.Worksheet xlWorksheet;

            xlAppl = new ExcelRef.Application();
            xlWorkbook = xlAppl.Workbooks.Add(Missing.Value);
            xlWorksheet = xlWorkbook.ActiveSheet;
           // xlAppl.Visible = true;

            CultureInfo attuale = Thread.CurrentThread.CurrentUICulture;
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("it-IT");

            // ------------------------------------
            // --------- Scrivere nelle colonne i mesi dell'anno e mettere a 7 la loro larghezza
            shiftColonna = 3;
            riga = 3;
            String meseCode;
            for (int index = 1; index <= totaleMesiAnno; index++)
            {
                meseCode = index.ToString();
                colonna = index + shiftColonna;
                xlWorksheet.Cells[riga, colonna] = CultureInfo.CurrentUICulture.DateTimeFormat.GetAbbreviatedMonthName(index) + ".";
                xlWorksheet.Cells[riga, colonna].Font.Bold = true;
                xlWorksheet.Columns[colonna].ColumnWidth = 7;
                borderDraw(xlWorksheet, riga, colonna);

                meseColonna.Add(meseCode, colonna);
                monthTotHours.Add(meseCode, 0.0);
            }
            meseColonna.Add("keyColTot", shiftColonna);

            // ------------------------------------
            // --------- Elencare nelle prime 2 colonne i cognomi e nomi
            OleDbConnection connection = ConnectionHelper.createConnection();
            try { connection.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryString = "SELECT Gev_id, Cognome, Nome, Ruolo FROM tab_DatiAnagrafici INNER JOIN tabd_Ruoli ";
            queryString = queryString + "ON tab_DatiAnagrafici.ruolo = tabd_Ruoli.Codice ";
            queryString = queryString + "WHERE NOT tabd_Ruoli.ToIgnore ";
            queryString = queryString + "ORDER BY Cognome, Nome ";

            OleDbCommand dbCommand = new OleDbCommand(queryString, connection);
            OleDbDataReader dataReader = dbCommand.ExecuteReader();
            Dictionary<int, int> gevRiga = new Dictionary<int, int>();
            shiftRiga = 4;
            colonna = 1;
            while (dataReader.Read())
            {
                totaleRecordGev++;
                riga = totaleRecordGev + shiftRiga;
                xlWorksheet.Cells[riga, colonna] = dataReader["Cognome"];
                xlWorksheet.Cells[riga, colonna + 1] = dataReader["Nome"];
                borderDraw(xlWorksheet, riga, colonna);
                borderDraw(xlWorksheet, riga, colonna + 1);

                gevRiga.Add(Convert.ToInt32(dataReader["Gev_id"]), riga);
            }
            xlWorksheet.Columns[colonna].ColumnWidth = 16;
            xlWorksheet.Columns[colonna + 1].ColumnWidth = 17;
            dataReader.Close();

            // ------------------------------------
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlInsideHorizontal).Weight = BordersWeightRef.xlHairline;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlInsideVertical).Weight = BordersWeightRef.xlHairline;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeTop).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeRight).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeBottom).Weight = BordersWeightRef.xlMedium;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeLeft).Weight = BordersWeightRef.xlMedium;

            xlWorksheet.Cells[1, 1].Value = "Prospetto ore mensili";
            xlWorksheet.Cells[1, 1].Font.Size = 24;
            xlWorksheet.Cells[2, 1].Value = fromDate;
            xlWorksheet.Cells[2, 1].HorizontalAlignment = ExcelRef.Constants.xlCenter;
            xlWorksheet.Cells[2, 1].Font.Size = 16;
            xlWorksheet.Cells[2, 2].Value = toDate;
            xlWorksheet.Cells[2, 2].HorizontalAlignment = ExcelRef.Constants.xlCenter;
            xlWorksheet.Cells[2, 2].Font.Size = 16;

            int currentGev;
            String currentMonth;
            double totOre, gevTotHours;
            Boolean singolo, lastSingolo;

            queryString = "SELECT tab_DatiAnagrafici.Gev_id AS GevID, tab_DatiAnagrafici.Cognome, datepart('m',tab_OrdiniDiServizio.[Data]) AS numMese, tab_OrdiniDiServizio.DURATA AS Ore ";
            queryString = queryString + "FROM tab_OrdiniDiServizio INNER JOIN (tab_rel_Gev_Ods INNER JOIN (tab_DatiAnagrafici INNER JOIN tabd_Ruoli ON tab_DatiAnagrafici.ruolo = tabd_Ruoli.Codice) "; 
            queryString = queryString + "ON tab_rel_Gev_Ods.Gev_ID = tab_DatiAnagrafici.Gev_ID ) ON tab_OrdiniDiServizio.ODS_ID = tab_rel_Gev_Ods.Ods_ID ";
            queryString = queryString + "WHERE tab_OrdiniDiServizio.Data BETWEEN  cDate('" + fromDate.ToShortDateString() + "') AND cDate('" + toDate.ToShortDateString() + "') ";
            queryString = queryString + "AND tab_rel_Gev_Ods.D_END = cDate('31/12/9999') ";
            queryString = queryString + "AND NOT tabd_Ruoli.ToIgnore ";
            queryString = queryString + "ORDER BY tab_DatiAnagrafici.Gev_id,  tab_OrdiniDiServizio.[Data]; ";

            dbCommand = new OleDbCommand(queryString, connection);
            dataReader = dbCommand.ExecuteReader();

            Boolean eof = !dataReader.Read();
            while (!eof) {
                currentGev = Convert.ToInt32(dataReader["GevID"]);
                currentMonth = dataReader["numMese"].ToString();
                totOre = ((DateTime)dataReader["Ore"]).ToOADate(); 
                singolo = true;
                lastSingolo = true;
                gevTotHours = 0.0;
                riga = gevRiga[currentGev];

                eof = !dataReader.Read();
                if (!eof) {
                    while (Convert.ToInt32(dataReader["GevID"]) == currentGev)
                    {
                    singolo = false;
                    while (dataReader["numMese"].ToString().Equals(currentMonth)
                            && Convert.ToInt32(dataReader["GevID"]) == currentGev)
                    {
                        totOre = totOre + ((DateTime)dataReader["Ore"]).ToOADate();
                        eof = !dataReader.Read();
                        lastSingolo = false;
                        if (eof) break;
                    }
                    // scrivi nel foglio
                    colonna = meseColonna[currentMonth];
                    xlWorksheet.Cells[riga, colonna].Value = GeneralHelper.ConvertDaysToHours(totOre);
                    xlWorksheet.Cells[riga, colonna].NumberFormat = "[h]:mm";
                    gevTotHours = gevTotHours + totOre;
                    monthTotHours[currentMonth] = monthTotHours[currentMonth] + totOre;

                    // passa al mese successivo
                    if (!eof){
                        if (Convert.ToInt32(dataReader["GevID"]) == currentGev)
                        {
                            currentMonth = dataReader["numMese"].ToString();
                            totOre = ((DateTime)dataReader["Ore"]).ToOADate();
                            eof = !dataReader.Read();
                            lastSingolo = true;
                        }
                    }
                    if (eof) break;
                  }
                }

                // caso di un solo ods
                if (singolo || lastSingolo) {
                    colonna = meseColonna[currentMonth];
                    xlWorksheet.Cells[riga, colonna].Value = GeneralHelper.ConvertDaysToHours(totOre);
                    xlWorksheet.Cells[riga, colonna].NumberFormat = "[h]:mm";
                    gevTotHours = gevTotHours + totOre;
                    monthTotHours[currentMonth] = monthTotHours[currentMonth] + totOre;
                }
                colonna = meseColonna["keyColTot"];
                xlWorksheet.Cells[riga, colonna].Value = GeneralHelper.ConvertDaysToHours(gevTotHours);
                xlWorksheet.Cells[riga, colonna].NumberFormat = "[h]:mm";
                xlWorksheet.Cells[riga, colonna].Font.ColorIndex = 3;
            }
            dataReader.Close();

            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, shiftColonna]].Borders(BordersRef.xlEdgeTop).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, shiftColonna]].Borders(BordersRef.xlEdgeBottom).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[1 + shiftRiga, shiftColonna], xlWorksheet.Cells[totaleRecordGev + shiftRiga, shiftColonna]].Borders(BordersRef.xlInsideHorizontal).LineStyle = LineStyleRef.xlContinuous;
 
          //And we can iterate through the complete dictionary
            int coco = shiftColonna + 1;
            double totaleAnnuo = 0;
            foreach(KeyValuePair<String, double> kvp in monthTotHours){
                xlWorksheet.Cells[shiftRiga, coco] = GeneralHelper.ConvertDaysToHours(kvp.Value);
                xlWorksheet.Cells[shiftRiga, coco].NumberFormat = "[h]:mm";
                xlWorksheet.Cells[shiftRiga, coco].Font.ColorIndex = 30;
                coco = coco + 1;
                totaleAnnuo = totaleAnnuo + kvp.Value;
            }
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeLeft).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlEdgeRight).LineStyle = LineStyleRef.xlContinuous;
            xlWorksheet.Range[xlWorksheet.Cells[shiftRiga, 1 + shiftColonna], xlWorksheet.Cells[shiftRiga, totaleMesiAnno + shiftColonna]].Borders(BordersRef.xlInsideVertical).LineStyle = LineStyleRef.xlContinuous;
 
            xlWorksheet.Cells[shiftRiga, shiftColonna] = GeneralHelper.ConvertDaysToHours(totaleAnnuo);
            xlWorksheet.Cells[shiftRiga, shiftColonna].Font.Color = ColorRef.rgbBlue;
            xlWorksheet.Cells[shiftRiga, shiftColonna].NumberFormat = "[h]:mm";

            xlAppl.Visible = true;

            Thread.CurrentThread.CurrentUICulture = attuale;
            releaseObject(xlWorksheet);
            releaseObject(xlWorkbook);
            releaseObject(xlAppl);
        }

        private void borderDraw(ExcelRef.Worksheet foglio, int ri, int co){
          foglio.Cells[ri, co].Borders(BordersRef.xlEdgeLeft).LineStyle = LineStyleRef.xlContinuous;
          foglio.Cells[ri, co].Borders(BordersRef.xlEdgeTop).LineStyle = LineStyleRef.xlContinuous;
          foglio.Cells[ri, co].Borders(BordersRef.xlEdgeRight).LineStyle = LineStyleRef.xlContinuous;
          foglio.Cells[ri, co].Borders(BordersRef.xlEdgeBottom).LineStyle = LineStyleRef.xlContinuous;
        }

        private int ElencaCognomiNomi(ExcelRef.Worksheet xlWorksheetP, Dictionary<int, int> gevRigaP, int shiftRigaP)
        {
            String queryStringL;
            int rigaL, colonnaL;
            int totaleRecordGevL = 0;

            OleDbConnection connectionL = ConnectionHelper.createConnection();
            try { connectionL.Open(); }
            catch (OleDbException ode) { MessageBox.Show("ATTENZIONE: " + ode); }

            queryStringL = "SELECT Gev_id, Cognome, Nome, Ruolo FROM tab_DatiAnagrafici INNER JOIN tabd_Ruoli ";
            queryStringL = queryStringL + "ON tab_DatiAnagrafici.ruolo = tabd_Ruoli.Codice ";
            queryStringL = queryStringL + "WHERE NOT tabd_Ruoli.ToIgnore ";
            queryStringL = queryStringL + "ORDER BY Cognome, Nome ";

            OleDbCommand dbCommand = new OleDbCommand(queryStringL, connectionL);
            OleDbDataReader dataReader = dbCommand.ExecuteReader();
            colonnaL = 1;
            while (dataReader.Read())
            {
                totaleRecordGevL++;
                rigaL = totaleRecordGevL + shiftRigaP;
                xlWorksheetP.Cells[rigaL, colonnaL] = dataReader["Cognome"];
                xlWorksheetP.Cells[rigaL, colonnaL + 1] = dataReader["Nome"];
                borderDraw(xlWorksheetP, rigaL, colonnaL);
                borderDraw(xlWorksheetP, rigaL, colonnaL + 1);

                gevRigaP.Add(Convert.ToInt32(dataReader["Gev_id"]), rigaL);
            }
            xlWorksheetP.Columns[colonnaL].ColumnWidth = 16;
            xlWorksheetP.Columns[colonnaL + 1].ColumnWidth = 17;
            dataReader.Close();

            return totaleRecordGevL;
        }

        public void excelListaServizi(DataGridView dataGridView, DateTimePicker from, DateTimePicker to, Boolean typeSelected)
        {
            int riga, colonna;
            String tipoSelezionato = "";
            
            ExcelRef.Application xlAppl;
            ExcelRef.Workbook xlWorkbook;
            ExcelRef.Worksheet xlWorksheet;

            xlAppl = new ExcelRef.Application();
            xlWorkbook = xlAppl.Workbooks.Add(Missing.Value);
            xlWorksheet = xlWorkbook.ActiveSheet;

            xlWorksheet.Columns[2].ColumnWidth = 12;
            xlWorksheet.Columns[3].ColumnWidth = 15;
            xlWorksheet.Columns[4].ColumnWidth = 10;
            xlWorksheet.Columns[5].ColumnWidth = 40;

            xlAppl.Visible = true;

            riga = 2;
            xlWorksheet.Cells[riga, 2] = "periodo dal: ";
            xlWorksheet.Cells[riga, 3] = from.Value.ToString("MM/dd/yyyy");
            xlWorksheet.Cells[riga, 3].Font.Size = 14;
            xlWorksheet.Cells[riga, 3].HorizontalAlignment = ExcelRef.Constants.xlLeft;
            xlWorksheet.Cells[riga, 4] = "al ";
            xlWorksheet.Cells[riga, 5] = to.Value.ToString("MM/dd/yyyy");
            xlWorksheet.Cells[riga, 5].Font.Size = 14;
            xlWorksheet.Cells[riga, 5].HorizontalAlignment = ExcelRef.Constants.xlLeft;

            riga = 4;
            if (!typeSelected)
            {
                xlWorksheet.Cells[riga, 1] = "tipo";
                xlWorksheet.Cells[riga, 1].Font.Bold = true;
            }
            xlWorksheet.Cells[riga, 2] = "effetuato il";
            xlWorksheet.Cells[riga, 2].Font.Bold = true;
            xlWorksheet.Cells[riga, 3] = "con num. ODS";
            xlWorksheet.Cells[riga, 3].Font.Bold = true;
            xlWorksheet.Cells[riga, 4] = "in zona";
            xlWorksheet.Cells[riga, 4].Font.Bold = true;
            xlWorksheet.Cells[riga, 5] = "partecipanti";
            xlWorksheet.Cells[riga, 5].Font.Bold = true;

            riga = 5;
            foreach (DataGridViewRow row in dataGridView.Rows) { 
              colonna = 1;
              tipoSelezionato = row.Cells[4].Value.ToString().PadLeft(3, '0');
              if (!typeSelected)
              {
                  // TIPO
                  colonna = 0;
                  xlWorksheet.Cells[riga, ++colonna] = "'" + row.Cells[4].Value.ToString().PadLeft(3, '0');
                  xlWorksheet.Cells[riga, colonna].HorizontalAlignment = ExcelRef.Constants.xlCenter;
              }
              // DATA
              xlWorksheet.Cells[riga, ++colonna] = ((DateTime)row.Cells[2].Value).ToString("MM/dd/yyyy");
              // NUMERO ODS
              xlWorksheet.Cells[riga, ++colonna] = "'" + GeneralHelper.formatNumOds(row.Cells[1].Value.ToString());
              xlWorksheet.Cells[riga, colonna].HorizontalAlignment = ExcelRef.Constants.xlCenter;
              // ZONA
              xlWorksheet.Cells[riga, ++colonna] = row.Cells[5].Value;
              xlWorksheet.Cells[riga, colonna].HorizontalAlignment = ExcelRef.Constants.xlLeft;
              // PARTECIPANTI
              xlWorksheet.Cells[riga, ++colonna] = row.Cells[6].Value;   
                
              riga++;
            }

            String titolo = "Elenco servizi";
            if (typeSelected) titolo = "Elenco servizi di tipo " + tipoSelezionato;
            if (typeSelected && tipoSelezionato.Equals("190")) titolo = "Elenco delle vigilanze ";
            if (typeSelected && tipoSelezionato.Equals("191")) titolo = "Elenco delle vigilanze con guardia parco";
            xlWorksheet.Cells[1, 2] = titolo;
            xlWorksheet.Cells[1, 2].Font.Size = 18;

            releaseObject(xlWorksheet);
            releaseObject(xlWorkbook);
            releaseObject(xlAppl);
        }


        private void releaseObject(object obj)
        {
          try
          {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
          }
          catch (Exception ex)
          {
            obj = null;
            MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
          }
          finally
          {
            GC.Collect();
          }
        }

    }
}
