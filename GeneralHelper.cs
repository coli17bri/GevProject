using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Drawing;

namespace WindowsFormsApplication1
{
    public static class GeneralHelper
    {
        private static double FIX_FACTOR = 0.2;
        private static Font SMALL_FONT = new System.Drawing.Font("Microsoft Sans Serif", 
                                                                  8.75F, 
                                                                  System.Drawing.FontStyle.Regular,
                                                                  System.Drawing.GraphicsUnit.Point, 
                                                                  ((byte)(0)));

        public static void createMulticolumnCombobox(ComboBox cbox, double factor)
        {
            // Enable the owner draw on the ComboBox.
            cbox.DrawMode = DrawMode.OwnerDrawFixed;
            // Handle the DrawItem event to draw the items.
            cbox.DrawItem += delegate(object cmb, DrawItemEventArgs args)
            {
                // Draw the default background
                args.DrawBackground();
                // The ComboBox is bound to a DataTable,
                // so the items are DataRowView objects.
                DataRowView drv = (DataRowView)cbox.Items[args.Index];
                // Retrieve the value of each column.
                string code = drv["Codice"].ToString();
                string description = drv["Descrizione"].ToString();

                // Get the bounds for the first column
                Rectangle r1 = args.Bounds;
                r1.Width = (int)(r1.Width * factor);
                // Draw the text on the first column
                using (SolidBrush sb = new SolidBrush(args.ForeColor))
                {
                    args.Graphics.DrawString(code, args.Font, sb, r1);
                }

                // Draw a line to isolate the columns 
                using (Pen p = new Pen(Color.Black))
                {
                    args.Graphics.DrawLine(p, r1.Right, 0, r1.Right, r1.Bottom);
                }

                // Get the bounds for the second column
                Rectangle r2 = args.Bounds;
                r2.X = (int)(args.Bounds.Width * factor);
                r2.Width = (int)(r2.Width * (1 - factor));

                // Draw the text on the second column
                using (SolidBrush sb = new SolidBrush(args.ForeColor))
                {
                    args.Graphics.DrawString(description, SMALL_FONT, sb, r2);
                }
            };

        }

        public static void createMulticolumnCombobox(ComboBox cbox)
        {
            createMulticolumnCombobox(cbox, FIX_FACTOR); 
        }

        public static String formatNumOds(String strInput) {
          String inputNum = strInput.Replace(".", ",");
          String outputNum = ((Convert.ToDouble(inputNum)).ToString("F1")).PadLeft(6, '0');
          return outputNum.Replace(",", ".");
        }

        public static String formatDurataOds(String strInput)
        {
            DateTime miaData;
            try  {
                miaData = Convert.ToDateTime(strInput);
                return miaData.ToShortTimeString();
            }
            catch (FormatException) {                
                return "04:00";
            }
        }

        public static string ConvertDaysToHours(double sommaOre)
        {
            String totHours = "err:00";
            int giorni = 0, giorniInOre;
            int soloOre, soloMinuti, oreTotali;

            try
            {
                // aggiungo un secondo per correggere il bug che considera 0.99999999 come 0 giorni
                double unSecondo = (double)(1.0 / 86400.0);
                sommaOre = sommaOre + unSecondo;
                DateTime dtSommaOre = DateTime.FromOADate(sommaOre);
                soloOre = dtSommaOre.Hour;
                soloMinuti = dtSommaOre.Minute;

                giorni = (int)Math.Floor(sommaOre);
                giorniInOre = giorni * 24;

                oreTotali = giorniInOre + soloOre;
                totHours = oreTotali.ToString().PadLeft(2, '0') + ":" + soloMinuti.ToString().PadLeft(2, '0');
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return totHours;
        }

    }
}
