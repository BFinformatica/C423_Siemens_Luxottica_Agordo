using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace GestioneDatabase
{
    public class Sad : Database
    {
        public Sad() { }
        public Sad(TipologiaConnessione connessione)
        {
            this.Connessione = connessione;
        }
        /// <summary>
        /// Restituisce un tatatable con il sad caricato
        /// </summary>
        /// <param name="sql">inutile per i file sad</param>
        /// <param name="parametri">gli eventuali parametri con cui filtrare il sad</param>
        /// <returns></returns>
        public override DataTable Reader(string sql, Parameter[] parametri = null)
        {
            System.IO.StreamReader leggi = new System.IO.StreamReader(this.Connessione.ConnectionStringCompleta);
            int num_riga = 0;
            DataTable tbl = new DataTable();
            int index = -1;
            while (!leggi.EndOfStream)
            {
                string riga = leggi.ReadLine();
                switch (num_riga)
                {
                    case 0:
                        //this.Software = riga.TrimStart().TrimEnd();
                        break;
                    case 1:
                        //this.Impianto = riga.TrimStart().TrimEnd();
                        break;
                    case 2:
                        foreach (var i in riga.Replace("#", "Date").Split('\t'))
                        {
                            if ((i.TrimEnd().TrimStart() == "") && (parametri != null) && (parametri[0].Name == "misura") && (((tbl.Columns[tbl.Columns.Count - 1].ColumnName.Split('_')[0] == parametri[0].Value.ToString()) && (!tbl.Columns[tbl.Columns.Count - 1].ColumnName.Contains("_val"))) || (tbl.Columns[tbl.Columns.Count - 1].ColumnName == "Date")))
                                tbl.Columns.Add(tbl.Columns[tbl.Columns.Count - 1].ColumnName + "_val");
                            else if ((parametri != null) && (parametri[0].Name == "misura") && ((i.TrimEnd().TrimStart().Split('_')[0] == parametri[0].Value.ToString()) || (i.TrimEnd().TrimStart() == "Date")))
                            {
                                tbl.Columns.Add(i.TrimEnd().TrimStart());
                                index = riga.Replace("#", "Date").Split('\t').ToList().IndexOf(i);
                            }
                            else if ((i.TrimEnd().TrimStart() == "") && (parametri == null))
                                tbl.Columns.Add(tbl.Columns[tbl.Columns.Count - 1].ColumnName + "_val");
                            else if (parametri == null)
                                tbl.Columns.Add(i.TrimEnd().TrimStart());
                        }
                        tbl.Columns["Date_val"].ColumnName = "Hour";
                        break;
                    case 3:
                        //tbl.Columns[2].ColumnName = tbl.Columns[2].ColumnName + " (" + riga.Split('\t')[index] + ")";
                        break;
                    default:
                        //this.Rows.Add(new RigaSad(intestazoini, unita_misura, riga.Split('\t')));
                        tbl.Rows.Add(tbl.NewRow());
                        tbl.Rows[tbl.Rows.Count - 1][0] = riga.Split('\t')[0];
                        tbl.Rows[tbl.Rows.Count - 1][1] = riga.Split('\t')[1];
                        tbl.Rows[tbl.Rows.Count - 1][2] = riga.Split('\t')[index];
                        break;
                }
                num_riga++;
            }
            leggi.Close();
            //sistemo i campi Date e Hour
            foreach (DataRow r in tbl.Rows)
                r["Date"] = this.SistemaData(r["Date"].ToString(), r["Hour"].ToString());
            tbl.Columns.Remove("Hour");
            tbl.Columns["Date"].ColumnName = "DateHour";
            tbl.Columns[1].ColumnName = "DT_VALUE";
            return tbl;
        }
        /// <summary>
        /// Restituisce un tatatable con il sad caricato
        /// </summary>
        /// <param name="sql">inutile per i file sad</param>
        /// <param name="parametri">gli eventuali parametri con cui filtrare il sad</param>
        /// <returns></returns>
        public override void Query(string sql, Parameter[] parametri = null)
        {
        }
        /// <summary>
        /// Inutile per il Sad
        /// </summary>
        public override void Open()
        {
            //nulla da fare
        }
        /// <summary>
        /// Inutile per il Sad
        /// </summary>
        public override void Close()
        {
            //nulla da fare
        }
        /// <summary>
        /// Cerca se la cartella che contiene i sad esiste
        /// </summary>
        /// <param name="dbname">Nome della cartella</param>
        /// <returns></returns>
        public override bool DatabaseExist(string dbname)
        {
            return System.IO.Directory.Exists(dbname);
        }
        /// <summary>
        /// Cerca se il file sad specificato esiste
        /// </summary>
        /// <param name="tableName">nome del file</param>
        /// <returns></returns>
        public override bool TableExist(string tableName)
        {
            return System.IO.File.Exists(this.Connessione.Server + "\\" + tableName);
        }
        /// <summary>
        /// Cerca se esiste un file sad
        /// </summary>
        /// <param name="table_filter">il valore per filtrare l'elenco dei file</param>
        /// <returns></returns>
        public override string[] GetTable(string table_filter)
        {
            return (from q in System.IO.Directory.GetFiles(this.Connessione.Server, "*.sad")
                    let nomeFile = System.IO.Path.GetFileNameWithoutExtension(q)
                    where nomeFile.Contains(table_filter)
                    select q).ToArray();

        }
        private DateTime SistemaData(string date, string time)
        {
            date = date.Insert(4, ".");
            date = date.Insert(7, ".");
            DateTime tempo = DateTime.Parse(date);
            tempo = tempo.AddHours(int.Parse(time.Split('.')[0]));
            tempo = tempo.AddMinutes(int.Parse(time.Split('.')[1]));
            tempo = tempo.AddSeconds(int.Parse(time.Split('.')[2]));
            return tempo;
        }
        public override string DatoGiorno(string tableName, string TipoAggregato, string stazione, string mese, string parametro, string campoDato, string CampoValidita, string validita, string StatoImpianto) { return ""; }
        public override string DatoMese(string tableName, string TipoAggregato, string stazione, string mese, string parametro, string campoDato, string CampoValidita, string validita, string StatoImpianto) { return ""; }
    }
}
