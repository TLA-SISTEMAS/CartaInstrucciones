using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CartaInstrucciones
{
    public partial class CartadeInstruccionesMenu : Form
    {

        private static CartadeInstruccionesMenu _instancia;

        public CartadeInstruccionesMenu()
        {
            InitializeComponent();
        }

        public static CartadeInstruccionesMenu Instancia
        {
            get
            {
                if (_instancia == null || _instancia.IsDisposed)
                {
                    _instancia = new CartadeInstruccionesMenu();
                }
                return _instancia;
            }
        }

        private void CartadeInstruccionesMenu_Load(object sender, EventArgs e)
        {

        }

        private void btnAltasCarta_Click(object sender, EventArgs e)
        {
            CartadeInstruccionesLlenado miCartadeInstruccionesLlena = CartadeInstruccionesLlenado.Instancia;

            Form mdiParent = this.MdiParent;

            if (miCartadeInstruccionesLlena.MdiParent == null)
            {
                miCartadeInstruccionesLlena.MdiParent = mdiParent;
            }

            if (miCartadeInstruccionesLlena.WindowState == FormWindowState.Minimized)
            {
                miCartadeInstruccionesLlena.WindowState = FormWindowState.Normal;
            }

            miCartadeInstruccionesLlena.Show(); // Show() se encarga de mostrarlo por primera vez
            miCartadeInstruccionesLlena.BringToFront();
            miCartadeInstruccionesLlena.Focus();
        }
    }
}
