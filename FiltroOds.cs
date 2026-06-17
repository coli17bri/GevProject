using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{

    class FiltroOds
    {

        private DateTime dataFrom = new DateTime(DateTime.Today.Year, 1, 1);
        private DateTime dataTo = new DateTime(DateTime.Today.Year, 12, 31);
        private String numeroFrom = "0000.0";
        private String numeroTo = "9999.9";
        private String tipo = "";
        private String luogo = "";

        public DateTime getDataFrom() { return dataFrom; }
        public void setDataFrom(DateTime pDataFrom) { dataFrom = pDataFrom; }

        public DateTime getDataTo() { return dataTo; }
        public void setDataTo(DateTime pDataTo) { dataTo = pDataTo; }

        public String getNumeroFrom() { return numeroFrom; }
        public void setNumeroFrom(String pNumeroFrom) { numeroFrom = pNumeroFrom; }

        public String getNumeroTo() { return numeroTo; }
        public void setNumeroTo(String pNumeroTo) { numeroTo = pNumeroTo; }

        public String getTipo() { return tipo; }
        public void setTipo(String pTipo) { tipo = pTipo; }

        public String getLuogo() { return luogo; }
        public void setLuogo(String pLuogo) { luogo = pLuogo; }
    }
}
