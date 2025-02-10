using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace GestioneDatabase
{
    public class Sql : Database
    {
        private SqlConnection conn;
        public Sql() { }
        public Sql(TipologiaConnessione connessione)
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
            SqlCommand query = new SqlCommand(sql, conn);
            if (parametri != null)
                for (int i = 0; i < parametri.Length; i++)
                    query.Parameters.Add(new SqlParameter(parametri[i].Name, parametri[i].Value));
            DataTable tbl = new DataTable();
            try
            {
                var read = query.ExecuteReader();
                tbl.Load(read);
                read.Close();
                read.Dispose();
            }
            catch (System.Exception err)
            {

            }
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
            SqlCommand query = new SqlCommand(sql, conn);
            if (parametri != null)
                for (int i = 0; i < parametri.Length; i++)
                    query.Parameters.Add(new SqlParameter(parametri[i].Name, parametri[i].Value));
            query.ExecuteNonQuery();
        }
        /// <summary>
        /// Apre la connessione al database
        /// </summary>
        public override void Open()
        {
            conn = new SqlConnection(this.Connessione.ConnectionStringCompleta);
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
            SqlConnection temp_conn = new SqlConnection(this.Connessione.ConnectionString);
            temp_conn.Open();
            SqlCommand query = new SqlCommand("SELECT name FROM master.dbo.sysdatabases WHERE(('[' + name + ']' = @dbname1) OR (name = @dbname2))", temp_conn);
            query.Parameters.Add(new SqlParameter("dbname1", dbname));
            query.Parameters.Add(new SqlParameter("dbname2", dbname));
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
            DataTable result = this.Reader("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @tableName", new Parameter[] { new Parameter("tableName", tableName) });
            return (result.Rows.Count > 0);
        }
        /// <summary>
        /// Cerca una tabella all'interno del database
        /// </summary>
        /// <param name="table_filter">filtro sul nome della tabella</param>
        /// <returns></returns>
        public override string[] GetTable(string table_filter)
        {
            DataTable result = this.Reader("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE @tableName", new Parameter[] { new Parameter("tableName", "%" + table_filter + "%") });
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
