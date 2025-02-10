using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;

namespace GestioneDatabase
{
    public class MysqlDb : Database
    {
        private MySqlConnection conn;
        public MysqlDb() { }
        public MysqlDb(TipologiaConnessione connessione)
        {
            this.Connessione = connessione;
            this.Open();
        }
        /// <summary>
        /// Restituisce un tatatable con i risultati
        /// </summary>
        /// <param name="sql">il codice sql</param>
        /// <param name="parametri">gli eventuali parametri</param>
        /// <returns></returns>
        public override DataTable Reader(string sql, Parameter[] parametri = null)
        {
            MySqlCommand query = new MySqlCommand(sql, conn);
            if (parametri != null)
                for (int i = 0; i < parametri.Length; i++)
                    query.Parameters.Add(new MySqlParameter(parametri[i].Name, parametri[i].Value));
            DataTable tbl = new DataTable();
            var read = query.ExecuteReader();
            tbl.Load(read);
            read.Close();
            read.Dispose();
            return tbl;
        }
        /// <summary>
        /// Esegue una query senza risultati
        /// </summary>
        /// <param name="sql">il codice sql</param>
        /// <param name="parametri">gli eventuali parametri</param>
        /// <returns></returns>
        public override void Query(string sql, Parameter[] parametri = null)
        {
            MySqlCommand query = new MySqlCommand(sql, conn);
            if (parametri != null)
                for (int i = 0; i < parametri.Length; i++)
                    query.Parameters.Add(new MySqlParameter(parametri[i].Name, parametri[i].Value));
            query.ExecuteNonQuery();
        }
        /// <summary>
        /// Apre la connessione al database
        /// </summary>
        public override void Open()
        {
            conn = new MySqlConnection(this.Connessione.ConnectionStringCompleta);
            conn.Open();
        }
        /// <summary>
        /// Chiude la connessione al database
        /// </summary>
        public override void Close()
        {
            conn.Close();
        }
        /// <summary>
        /// Cerca se il database esiste
        /// </summary>
        /// <param name="dbname">Nome del database</param>
        /// <returns></returns>
        public override bool DatabaseExist(string dbname)
        {
            MySqlConnection temp_conn = new MySqlConnection(this.Connessione.ConnectionString);
            temp_conn.Open();
            MySqlCommand query = new MySqlCommand("SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = @dbname", temp_conn);
            query.Parameters.Add(new MySqlParameter("dbname", dbname));
            DataTable tbl = new DataTable();
            var read = query.ExecuteReader();
            tbl.Load(read);
            read.Close();
            temp_conn.Close();
            return (tbl.Rows.Count > 0);
        }
        /// <summary>
        /// Cerca se la tabella esiste
        /// </summary>
        /// <param name="tableName">nome della tabella</param>
        /// <returns></returns>
        public override bool TableExist(string tableName)
        {
            DataTable result = this.Reader("SELECT * FROM information_schema.tables WHERE (table_name = @tableName)", new Parameter[] { new Parameter("tableName", tableName) });
            return (result.Rows.Count > 0);
        }
        /// <summary>
        /// Cerca una tabella all'interno del database
        /// </summary>
        /// <param name="table_filter">filtro sul nome della tabella</param>
        /// <returns></returns>
        public override string[] GetTable(string table_filter)
        {
            DataTable result = this.Reader("SELECT * FROM information_schema.tables WHERE (table_name LIKE @tableName)", new Parameter[] { new Parameter("tableName", "%" + table_filter + "%") });
            return (from q in result.AsEnumerable() select q["TABLE_NAME"].ToString()).ToArray();
        }
        public override string DatoGiorno(string tableName, string TipoAggregato, string stazione, string giorno, string parametro, string campoDato, string CampoValidita, string validita, string StatoImpianto)
        {
            var tabella = this.Reader("SELECT " + TipoAggregato + " (" + campoDato + ") AS dato FROM " + tableName + " WHERE ((DT_STATIONCODE = @stazione) AND (DT_MeasureCod = @parametro) AND (DT_DATE LIKE @giorno) AND (DT_Custom1 = "
                + StatoImpianto + ")) GROUP by dt_date", new Parameter[]
                {
                    new Parameter("stazione", stazione),
                    new Parameter("parametro", parametro),
                    new Parameter("mese", giorno)
                });
            if (CampoValidita != "")
                tabella = this.Reader("SELECT " + TipoAggregato + " (" + campoDato + ") AS dato FROM " + tableName + " WHERE ((DT_STATIONCODE = @stazione) AND (DT_MeasureCod = @parametro) AND (DT_DATE LIKE @giorno) AND ("
                    + CampoValidita + " IN(@validita)) AND (DT_Custom1 = " + StatoImpianto + ")) GROUP by dt_date", new Parameter[]
                    {
                        new Parameter("stazione", stazione),
                        new Parameter("parametro", parametro),
                        new Parameter("mese", giorno),
                        new Parameter("validita", validita)
                    });
            return tabella.Rows[0][TipoAggregato].ToString();
        }
        public override string DatoMese(string tableName, string TipoAggregato, string stazione, string mese, string parametro, string campoDato, string CampoValidita, string validita, string StatoImpianto)
        {
            var tabella = this.Reader("SELECT " + TipoAggregato + " (" + campoDato + ") AS dato FROM " + tableName + " WHERE ((DT_STATIONCODE = @stazione) AND (DT_MeasureCod = @parametro) AND (DT_DATE LIKE @mese) AND (DT_Custom1 = "
                + StatoImpianto + "))", new Parameter[]
                {
                    new Parameter("stazione", stazione),
                    new Parameter("parametro", parametro),
                    new Parameter("mese", mese)
                });
            if (CampoValidita != "")
                tabella = this.Reader("SELECT " + TipoAggregato + " (" + campoDato + ") AS dato FROM " + tableName + " WHERE ((DT_STATIONCODE = @stazione) AND (DT_MeasureCod = @parametro) AND (DT_DATE LIKE @mese) AND ("
                    + CampoValidita + " IN(@validita)) AND (DT_Custom1 = " + StatoImpianto + "))", new Parameter[]
                    {
                        new Parameter("stazione", stazione),
                        new Parameter("parametro", parametro),
                        new Parameter("mese", mese),
                        new Parameter("validita", validita)
                    });
            return tabella.Rows[0][TipoAggregato].ToString();
        }
    }
}
