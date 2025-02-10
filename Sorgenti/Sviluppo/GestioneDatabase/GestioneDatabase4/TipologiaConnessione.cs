using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GestioneDatabase
{
    public class TipologiaConnessione
    {
        public TipologiaConnessione(string _server, string _database, string _username, string _password, string _selectedstationcode, Tipo t = Tipo.SqlServer)
        {
            this.tipo = t;
            this.Server = _server;
            this.Database = _database;
            this.Username = _username;
            this.Password = _password;
            this.SelectedStationCode = _selectedstationcode;
        }
        public TipologiaConnessione(System.Data.DataRow riga)
        {
            switch (riga["DBTYPE"].ToString())
            {
                case "SQL":
                    this.tipo = Tipo.SqlServer;
                    break;
                case "MYSQL":
                    this.tipo = Tipo.MySql;
                    break;
                default: break;
            }
            this.Server = riga["SERVER"].ToString();
            this.Database = riga["DATABASE"].ToString();
            this.Username = this.Decrypta(riga["USER"].ToString());
            this.Password = this.Decrypta(riga["PASSWORD"].ToString());
            this.SelectedStationCode = riga["STATIONCODE"].ToString();
            this.IsDefault = bool.Parse(riga["DEFAULT"].ToString());
            this.Description = riga["Description"].ToString();
        }
        public string ConnectionStringCompleta { get { return this.getConnedctionStringCompleta(); } }
        public string ConnectionString { get { return this.getConnedctionString(); } }
        public string Server { get; set; }
        public string Database { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public Tipo tipo { get; set; }
        public string SelectedStationCode { get; set; }
        public bool IsDefault { get; set; }
        public string Description { get; set; }

        #region Metodi Pubblici

        private string getConnedctionStringCompleta()
        {
            switch (this.tipo)
            {
                case Tipo.SqlServer: return "Server=" + this.Server + ";Database=" + this.Database + ";User Id=" + this.Username + ";Password=" + this.Password;
                case Tipo.MySql:
                    if (this.Database == "")
                        return "Server=" + this.Server + ";Uid=" + this.Username + ";Pwd=" + this.Password + ";SslMode=none";
                    return "Server=" + this.Server + ";Database=" + this.Database + ";Uid=" + this.Username + ";Pwd=" + this.Password + ";SslMode=none";
                case Tipo.Sad: return this.Server + "\\" + this.Database;
                default: return "";
            }
        }
        private string getConnedctionString()
        {
            switch (this.tipo)
            {
                case Tipo.SqlServer: return "Server=" + this.Server + ";User Id=" + this.Username + ";Password=" + this.Password;
                case Tipo.MySql: return "Server=" + this.Server + ";Uid=" + this.Username + ";Pwd=" + this.Password;
                case Tipo.Sad: return this.Server;
                default: return "";
            }
        }

        #endregion

        #region Metodi privati

        private string Decrypta(string valore, string chiave = "attimo")
        {
            string dest = "";
            int offset = Convert.ToInt16(valore.Substring(0, 2), 16);
            int SrcAsc;
            int TmpSrcAsc;
            int KeyPos = 0;
            for (int i = 3; (i <= valore.Length); i = (i + 2))
            {
                SrcAsc = Convert.ToInt16((valore.Substring((i - 1), 2).Trim()), 16);
                TmpSrcAsc = SrcAsc;
                if (KeyPos < chiave.Length)
                    KeyPos++;
                else
                    KeyPos = 1;
                TmpSrcAsc = SrcAsc ^ chiave[KeyPos - 1];
                if ((TmpSrcAsc <= offset))
                    TmpSrcAsc = (255 + (TmpSrcAsc - offset));
                else
                    TmpSrcAsc = (TmpSrcAsc - offset);
                dest = (dest + ((char)(TmpSrcAsc)));
                offset = SrcAsc;
            }
            return dest;
        }

        #endregion
    }
}
