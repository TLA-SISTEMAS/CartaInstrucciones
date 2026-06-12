using CartaInstrucciones.Modelo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronWord;
using objWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Web;

namespace CartaInstrucciones
{
    /// <summary>
    /// Comentario
    /// </summary>
    public partial class Form1 : Form
    {
        LinkLabel linkVisitedLabel = new LinkLabel();
        int cantidadFacturasCapturarGlobal = 0;
        public Form1()
        {
            InitializeComponent();
            

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void tsmCatalogo_Click(object sender, EventArgs e)
        {

        }
    }
}
