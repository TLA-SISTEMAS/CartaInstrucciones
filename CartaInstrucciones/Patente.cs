using System;
using System.Collections.Generic;

using System.Data.SQLite;

using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using System.Data.SqlClient;

namespace CartaInstrucciones.Modelo
{
    internal class Patente
    {
        private static string cadena = ConfigurationManager.ConnectionStrings["cadena"].ConnectionString;

        private static Patente _intancia = null;

        public Patente() { }
        public static Patente Instancia
        {
            get
            {
                if (_intancia == null)
                {
                    _intancia = new Patente();
                }

                return _intancia;
            }
        }

        public int idPatente { get; set; }
        public string nombreAgente { get; set; }

        public List<Patente> ListarPatente()
        {
            List<Patente> oLista = new List<Patente>();
            using (SQLiteConnection conexion = new SQLiteConnection(cadena))
            {
                conexion.Open();
                string query = "SELECT IDPatente FROM patente;";
                SQLiteCommand cmd = new SQLiteCommand(query, conexion);
                cmd.CommandType = System.Data.CommandType.Text;
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        oLista.Add(new Patente()
                        {
                            idPatente = int.Parse(dr["IDPatente"].ToString()),
                        });
                    }
                }
            }
            return oLista;
        }

        public Patente ListarunaPatente(string idPatente)
        {

            Patente proveedor = new Patente();
            using (SQLiteConnection conexion = new SQLiteConnection(cadena))
            {
                conexion.Open();

                string query = "SELECT * FROM patente where IDPatente = @idPatente;";
                SQLiteCommand cmd = new SQLiteCommand(query, conexion);
                cmd.Parameters.Add(new SQLiteParameter("@idPatente", idPatente));
                cmd.CommandType = System.Data.CommandType.Text;
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        proveedor = new Patente()
                        {
                            idPatente = int.Parse(dr["IDPatente"].ToString()),
                            nombreAgente = dr["NombreAgente"].ToString(),

                        };
                    }
                }
            }
            return proveedor;
        }
    }
}