using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
namespace CartaInstrucciones.Modelo
{
    public class Proveedor
    {
        public int ID { get; set; }
        public string Nombre { get; set; }
        public string Direccion { get; set; }
        public string Taxid { get; set; }


        private static string cadena = ConfigurationManager.ConnectionStrings["cadena"].ConnectionString;

        private static Proveedor _instancia;

        public Proveedor() { }

        public static Proveedor Instancia
        {
            get
            {
                if (_instancia == null)
                {
                    _instancia = new Proveedor();
                }

                return _instancia;
            }
        }

        public bool Guardar(Proveedor proveedor)
        {
            bool respuesta = true;

            using (SQLiteConnection con = new SQLiteConnection(cadena))
            {
                con.Open();
                string query = "insert into proveedores(Nombre,Direccion,Taxid) values (@nombre,@direccion,@taxid)";
                SQLiteCommand cmd = new SQLiteCommand(query, con);
                cmd.Parameters.Add(new SQLiteParameter("@nombre", proveedor.Nombre));
                cmd.Parameters.Add(new SQLiteParameter("@direccion", proveedor.Direccion));
                cmd.Parameters.Add(new SQLiteParameter("@taxid", proveedor.Taxid));
                if (cmd.ExecuteNonQuery() < 1)
                {
                    respuesta = false;
                }
            }
            return respuesta;
        }

        public bool Cambios(Proveedor proveedor)
        {
            bool respuesta = true;

            using (SQLiteConnection con = new SQLiteConnection(cadena))
            {
                con.Open();
                string query = "update proveedores set Nombre = @nombre ,Direccion = @direccion,Taxid = @taxid where ID = @id";
                SQLiteCommand cmd = new SQLiteCommand(query, con);
                cmd.Parameters.Add(new SQLiteParameter("@id",proveedor.ID));
                cmd.Parameters.Add(new SQLiteParameter("@nombre", proveedor.Nombre));
                cmd.Parameters.Add(new SQLiteParameter("@direccion", proveedor.Direccion));
                cmd.Parameters.Add(new SQLiteParameter("@taxid", proveedor.Taxid));

                if (cmd.ExecuteNonQuery() < 1)
                {
                    respuesta = false;
                }
            }
            return respuesta;
        }

        public bool Eliminar(Proveedor proveedor)
        {
            bool respuesta = true;

            using (SQLiteConnection con = new SQLiteConnection(cadena))
            {
                con.Open();
                string query = "delete from proveedores where ID = @id";
                SQLiteCommand cmd = new SQLiteCommand(query, con);
                cmd.Parameters.Add(new SQLiteParameter("@id", proveedor.ID));

                if (cmd.ExecuteNonQuery() < 1)
                {
                    respuesta = false;
                }
            }
            return respuesta;
        }

        public List<Proveedor> Listar()
        {
            List<Proveedor> listaProveedor = new List<Proveedor>();
            using (SQLiteConnection conexion = new SQLiteConnection(cadena))
            {
                conexion.Open();
                string query = "SELECT * FROM proveedores;";
                SQLiteCommand cmd = new SQLiteCommand(query, conexion);
                cmd.CommandType = System.Data.CommandType.Text;
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        listaProveedor.Add(new Proveedor()
                        {
                            ID = int.Parse(dr["ID"].ToString()),
                            Nombre = dr["Nombre"].ToString(),
                            Direccion = dr["Direccion"].ToString(),
                            Taxid = dr["Taxid"].ToString(),
                        });
                    }
                }
            }
            return listaProveedor;
        }

        public Proveedor ListarunProvedor(int idProveedor)
        {

            Proveedor proveedor = new Proveedor();
            using (SQLiteConnection conexion = new SQLiteConnection(cadena))
            {
                conexion.Open();
                
                string query = "SELECT * FROM proveedores where ID = @idproveedor;";
                SQLiteCommand cmd = new SQLiteCommand(query, conexion);
                cmd.Parameters.Add(new SQLiteParameter("@idproveedor", idProveedor));
                cmd.CommandType = System.Data.CommandType.Text;
                using (SQLiteDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        proveedor = new Proveedor()
                        {
                            ID = int.Parse(dr["ID"].ToString()),
                            Direccion = dr["Direccion"].ToString(),
                            Nombre = dr["Nombre"].ToString(),
                            Taxid = dr["Taxid"].ToString()
                        };
                    }
                }
            }
            return proveedor;
        }

    }
}
