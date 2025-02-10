using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace GestioneDatabase
{
    public class Gestione
    {
        private TipologiaConnessione conn;
        private string path;
        public Gestione(Tipo tipo)
        {
            switch (tipo)
            {
                case Tipo.SqlServer:
                    this.Database = new Sql();
                    break;
                case Tipo.MySql:
                    this.Database = new MysqlDb();
                    break;
                case Tipo.Sad:
                    this.Database = new Sad();
                    break;
                default: break;
            }
        }
        public Gestione(string path_connessioni_xml)
        {
            path = path_connessioni_xml;
            conn = (from q in this.Connections where q.IsDefault select q).FirstOrDefault();
            switch (conn.tipo)
            {
                case Tipo.SqlServer:
                    this.Database = new Sql(conn);
                    break;
                case Tipo.MySql:
                    this.Database = new MysqlDb(conn);
                    break;
                default: break;
            }
        }
        public Gestione(string path_connessioni_xml, string descrizione)
        {
            path = path_connessioni_xml;
            conn = (from q in this.Connections where q.Description.ToUpper().TrimStart().TrimEnd() == descrizione.ToUpper().TrimStart().TrimEnd() select q).FirstOrDefault();
            switch (conn.tipo)
            {
                case Tipo.SqlServer:
                    this.Database = new Sql(conn);
                    break;
                case Tipo.MySql:
                    this.Database = new MysqlDb(conn);
                    break;
                default: break;
            }
        }

        #region Proprietà

        public Database Database { get; set; }

        #endregion

        #region Metodi Pubblici

        public TipologiaConnessione[] Connections
        {
            get
            {
                DataSet dataset = new DataSet();
                dataset.ReadXml(path);
                return (from q in dataset.Tables["row"].AsEnumerable() select new TipologiaConnessione(q)).ToArray();
            }
        }

        #endregion
    }
}
