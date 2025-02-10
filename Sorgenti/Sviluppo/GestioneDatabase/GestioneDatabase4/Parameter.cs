using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GestioneDatabase
{
    public class Parameter
    {
        public Parameter(string _nome, object _value)
        {
            this.Name = _nome;
            this.Value = _value;
        }
        public string Name { get; set; }
        public object Value { get; set; }
        public System.Data.SqlClient.SqlParameter getSqlParameter()
        {
            return new System.Data.SqlClient.SqlParameter(this.Name, this.Value);
        }
        public MySql.Data.MySqlClient.MySqlParameter getMySqlParameter()
        {
            return new MySql.Data.MySqlClient.MySqlParameter(this.Name, this.Value);
        }
    }
}
