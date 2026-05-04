using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CartaInstrucciones
{
    internal static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        /// 
        public static Form1 cartaInstrucciones;
        //public static DatosImportador datosImportador;
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(cartaInstrucciones = new Form1());
            //Application.Run(datosImportador = new DatosImportador());
            //datosImportador.Hide();
        }
    }
}
