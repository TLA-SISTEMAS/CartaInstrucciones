using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data.SqlClient;
using System.Collections;

namespace CartaInstrucciones.Modelo
{
    internal class Aduana
    {
        private static string cadena = ConfigurationManager.ConnectionStrings["cadena"].ConnectionString;

        private static Aduana _intancia = null;
        public int IDAduana { get; set; }
        public string Nombre { get; set; }

        public Aduana() { }

        public static Aduana Instancia
        {
            get
            {
                if (_intancia == null)
                {
                    _intancia = new Aduana();
                }

                return _intancia;
            }
        }

        public List<Aduana> ListarAduana(string patente)
        {
            List<Aduana> listaduanas = new List<Aduana>();

            using(SQLiteConnection con = new SQLiteConnection(cadena))
            {
                con.Open();
                string query = "SELECT aduana.IDAduana FROM aduana INNER JOIN aduanapatente ON aduana.IDAduana = aduanapatente.IDAduana WHERE aduanapatente.IDPatente = @patente;";

                SQLiteCommand cmd = new SQLiteCommand(query, con);
                cmd.Parameters.Add(new SQLiteParameter("@patente", patente));
                cmd.CommandType = System.Data.CommandType.Text;

                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        listaduanas.Add(new Aduana()
                        {
                            IDAduana = int.Parse(dr["IDAduana"].ToString()),
                        });
                    }
                }
            }

            return listaduanas;
        }

        public Aduana ListarunaAduana(string idAduana)
        {

            Aduana proveedor = new Aduana();
            using (SQLiteConnection conexion = new SQLiteConnection(cadena))
            {
                conexion.Open();

                string query = "SELECT * FROM aduana where IDAduana = @idAduana;";
                SQLiteCommand cmd = new SQLiteCommand(query, conexion);
                cmd.Parameters.Add(new SQLiteParameter("@idAduana", idAduana));
                cmd.CommandType = System.Data.CommandType.Text;
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        proveedor = new Aduana()
                        {
                            IDAduana = int.Parse(dr["IDAduana"].ToString()),
                            Nombre = dr["Nombre"].ToString(),

                        };
                    }
                }
            }
            return proveedor;
        }
    }
}
