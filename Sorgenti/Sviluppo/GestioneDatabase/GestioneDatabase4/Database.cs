using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GestioneDatabase
{
    public abstract class Database
    {
        /// <summary>
        /// Restituisce un tatatable con i risultati
        /// </summary>
        /// <param name="sql">il codice sql</param>
        /// <param name="parametri">gli eventuali parametri</param>
        /// <returns></returns>
        public abstract System.Data.DataTable Reader(string sql, Parameter[] parametri = null);
        public abstract void Query(string sql, Parameter[] parametri = null);
        public abstract void Open();
        public abstract void Close();
        public abstract bool DatabaseExist(string dbname);
        public abstract bool TableExist(string tableName);
        public abstract string[] GetTable(string table_filter);
        public abstract string DatoGiorno(string tableName, string TipoAggregato, string stazione, string giorno, string parametro, string campoDato, string CampoValidita, string validita, string StatoImpianto);
        public abstract string DatoMese(string tableName, string TipoAggregato, string stazione, string mese, string parametro, string campoDato, string CampoValidita, string validita, string StatoImpianto);
        public TipologiaConnessione Connessione { get; set; }
    }
}
