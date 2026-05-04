using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


namespace CartaInstrucciones
{
    internal class Importador
    {
        private static string cadena = ConfigurationManager.ConnectionStrings["cadena"].ConnectionString;

        private static Importador _instancia;

        public Importador() { }

        public static Importador Instancia
        {
            get
            {
                if (_instancia == null)
                {
                    _instancia = new Importador();
                }

                return _instancia;
            }
        }

        public int idImportador { get; set; }
        public string nombreImportador { get; set; }
        public string domicilio { get; set; }
        public string rfc { get; set; }
        public string representanteLegal { get; set; }
        public string rfcRepresentante { get; set; }

        public bool Guardar(Importador importador)
        {
            bool respuesta = true;

            using (SQLiteConnection con = new SQLiteConnection(cadena))
            {
                con.Open();
                string query = "insert into importador(NombreImportador,Domicilio,RFC,RepresentanteLegal,RFCRepresentante) values (@nombreimportador,@domicilio,@rfc,@representantelegal,@rfcrepresentante)";
                SQLiteCommand cmd = new SQLiteCommand(query, con);
                cmd.Parameters.Add(new SQLiteParameter("@nombreimportador", importador.nombreImportador));
                cmd.Parameters.Add(new SQLiteParameter("@domicilio", importador.domicilio));
                cmd.Parameters.Add(new SQLiteParameter("@rfc", importador.rfc));
                cmd.Parameters.Add(new SQLiteParameter("@representantelegal", importador.representanteLegal));
                cmd.Parameters.Add(new SQLiteParameter("@rfcrepresentante", importador.rfcRepresentante));

                if (cmd.ExecuteNonQuery() < 1)
                {
                    respuesta = false;
                }
            }
            return respuesta;
        }

        public bool Cambios(Importador importador)
        {
            bool respuesta = true;

            using (SQLiteConnection con = new SQLiteConnection(cadena))
            {
                con.Open();
                string query = "update importador set NombreImportador = @nombreimportador ,Domicilio = @domicilio,RFC = @rfc,RepresentanteLegal = @representantelegal,RFCRepresentante = @rfcrepresentante where IDImportador = @idImportador";
                SQLiteCommand cmd = new SQLiteCommand(query, con);
                cmd.Parameters.Add(new SQLiteParameter("@idImportador", importador.idImportador));
                cmd.Parameters.Add(new SQLiteParameter("@nombreimportador", importador.nombreImportador));
                cmd.Parameters.Add(new SQLiteParameter("@domicilio", importador.domicilio));
                cmd.Parameters.Add(new SQLiteParameter("@rfc", importador.rfc));
                cmd.Parameters.Add(new SQLiteParameter("@representantelegal", importador.representanteLegal));
                cmd.Parameters.Add(new SQLiteParameter("@rfcrepresentante", importador.rfcRepresentante));

                if (cmd.ExecuteNonQuery() < 1)
                {
                    respuesta = false;
                }
            }
            return respuesta;
        }

        public bool Eliminar(Importador importador)
        {
            bool respuesta = true;

            using (SQLiteConnection con = new SQLiteConnection(cadena))
            {
                con.Open();
                string query = "delete from importador where IDImportador = @idImportador";
                SQLiteCommand cmd = new SQLiteCommand(query, con);
                cmd.Parameters.Add(new SQLiteParameter("@idImportador", importador.idImportador));

                if (cmd.ExecuteNonQuery() < 1)
                {
                    respuesta = false;
                }
            }
            return respuesta;
        }

        public List<Importador> Listar()
        {
            List<Importador> listaImportadores = new List<Importador>();
            using (SQLiteConnection conexion = new SQLiteConnection(cadena))
            {
                conexion.Open();
                string query = "SELECT * FROM importador;";
                SQLiteCommand cmd = new SQLiteCommand(query, conexion);
                cmd.CommandType = System.Data.CommandType.Text;
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        listaImportadores.Add(new Importador()
                        {
                            idImportador = int.Parse(dr["IDImportador"].ToString()),
                            nombreImportador = dr["NombreImportador"].ToString(),
                            domicilio = dr["Domicilio"].ToString(),
                            rfc = dr["RFC"].ToString(),
                            representanteLegal = dr["RepresentanteLegal"].ToString(),
                            rfcRepresentante = dr["RFCRepresentante"].ToString(),
                        });
                    }
                }
            }
            return listaImportadores;
        }

        public List<Importador> ListarunImportador(int idImportador)
        {
            List<Importador> importador = new List<Importador>();

            using (SQLiteConnection conexion = new SQLiteConnection(cadena))
            {
                conexion.Open();
                string query = "SELECT * FROM importador where IDImportador = @idimportador;";
                SQLiteCommand cmd = new SQLiteCommand(query, conexion);
                cmd.Parameters.Add(new SQLiteParameter("@idimportador", idImportador));
                cmd.CommandType = System.Data.CommandType.Text;
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        importador.Add(new Importador()
                        {
                            idImportador = int.Parse(dr["IDImportador"].ToString()),
                            nombreImportador = dr["NombreImportador"].ToString(),
                            domicilio = dr["Domicilio"].ToString(),
                            rfc = dr["RFC"].ToString(),
                            representanteLegal = dr["RepresentanteLegal"].ToString(),
                            rfcRepresentante = dr["RFCRepresentante"].ToString(),
                        });
                    }
                }
            }

            return importador;
        }
    }
}
