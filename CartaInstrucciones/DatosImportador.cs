using CartaInstrucciones.Modelo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CartaInstrucciones
{
    public partial class DatosImportador : Form
    {
        private static DatosImportador _instancia;

        LinkLabel linkVisitedLabel = new LinkLabel();
        DataGridViewRow filaImportador = new DataGridViewRow();
        bool btnCambiosClic = false;

        public DatosImportador()
        {
            InitializeComponent();
            linkVisitedLabel = llbDatosImportador;
        }

        public static DatosImportador Instancia
        {
            get 
            { 
                if(_instancia == null || _instancia.IsDisposed)
                {
                    _instancia = new DatosImportador();
                }
                return _instancia;
            }
        }

        void DatosImportador_Load(object sender, EventArgs e)
        {
            dgvDatosImportador.DataSource = null;
            dgvDatosImportador.DataSource = Importador.Instancia.Listar();

        }

        private void llbCartaInstrucciones_MouseEnter(object sender, EventArgs e)
        {
            if (!llbCartaInstrucciones.LinkVisited) llbCartaInstrucciones.LinkColor = Color.Blue;
        }

        private void llbCartaInstrucciones_MouseLeave(object sender, EventArgs e)
        {
            llbCartaInstrucciones.LinkColor = Color.Black;
        }

        private void llbDatosImportador_MouseEnter(object sender, EventArgs e)
        {
            if (!llbDatosImportador.LinkVisited) llbDatosImportador.LinkColor = Color.Blue;
        }

        private void llbDatosImportador_MouseLeave(object sender, EventArgs e)
        {
            llbDatosImportador.LinkColor = Color.Black;
        }

        private void llbDatosProveedor_MouseEnter(object sender, EventArgs e)
        {
            if (!llbDatosProveedor.LinkVisited) llbDatosProveedor.LinkColor = Color.Blue;
        }

        private void llbDatosProveedor_MouseLeave(object sender, EventArgs e)
        {
            llbDatosProveedor.LinkColor = Color.Black;
        }



        private void llbCartaInstrucciones_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            Program.cartaInstrucciones.Show();    
            
        }


        private void llbDatosProveedor_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();

            DatosProveedor.Instancia.Show(); 
        }


        private void DatosImportador_FormClosed_1(object sender, FormClosedEventArgs e)
        {
            Program.cartaInstrucciones.Close();
        }

        private void btnAltasDatos_Click(object sender, EventArgs e)
        {
            gboxImportador.Enabled = true;
            btnCambiosDatos.Enabled = false;
            btnEliminarDatosImportador.Enabled = false;
            btnGuardarDatosImportador.Enabled = true;
            btnCambiosClic = false;
        }

        private void btnGuardarDatosImportador_Click(object sender, EventArgs e)
        {

            bool respuesta;
            if (!btnCambiosClic)
            {
                Importador miImportador = new Importador()
                {
                    nombreImportador = txtboxImportadorDatos.Text.Trim(),
                    domicilio = txtboxDomicilioFiscalDatos.Text.Trim(),
                    rfc = txtboxRFCImportadorDatos.Text.Trim(),
                    representanteLegal = txtboxRepresentanteLegalDatos.Text.Trim(),
                    rfcRepresentante = txtboxRFCRepresentanteDatos.Text.Trim()
                };
                if (!comprobrarVacios()) {
                    btnAltasDatos.Enabled = true;
                    btnGuardarDatosImportador.Enabled = false;
                    btnCambiosDatos.Enabled = true;
                    btnEliminarDatosImportador.Enabled = true;
                    gboxImportador.Enabled = false;
                    respuesta = Importador.Instancia.Guardar(miImportador);
                    if (respuesta) { mostarImportadores(); }
                    limpiarTextBox();
                }
                               
            }
            else
            {
                Importador miImportador = new Importador()
                {
                    idImportador = int.Parse(filaImportador.Cells[0].Value.ToString()),
                    nombreImportador = txtboxImportadorDatos.Text.Trim(),
                    domicilio = txtboxDomicilioFiscalDatos.Text.Trim(),
                    rfc = txtboxRFCImportadorDatos.Text.Trim(),
                    representanteLegal = txtboxRepresentanteLegalDatos.Text.Trim(),
                    rfcRepresentante = txtboxRFCRepresentanteDatos.Text.Trim()
                };               
                if (!comprobrarVacios()) {
                    btnAltasDatos.Enabled = true;
                    btnGuardarDatosImportador.Enabled = false;
                    btnCambiosClic = false;
                    btnEliminarDatosImportador.Enabled = true;
                    gboxImportador.Enabled = true;
                    respuesta = Importador.Instancia.Cambios(miImportador);
                    if (respuesta) { mostarImportadores(); }
                    limpiarTextBox();
                }
                filaImportador = new DataGridViewRow();
            }
            
        }

        public void limpiarTextBox()
        {
            //LIMPIAR TEXTBOX
            txtboxImportadorDatos.Text = null;
            txtboxDomicilioFiscalDatos.Text = null;
            txtboxRFCImportadorDatos.Text = null;
            txtboxRepresentanteLegalDatos.Text = null;
            txtboxRFCRepresentanteDatos.Text = null;
        }

        public void mostarImportadores()
        {
            dgvDatosImportador.DataSource = null;
            dgvDatosImportador.DataSource = Importador.Instancia.Listar();
        }

        private void btnCambiosDatos_Click(object sender, EventArgs e)
        {
            try
            {                
                txtboxImportadorDatos.Text = filaImportador.Cells[1].Value.ToString();
                txtboxDomicilioFiscalDatos.Text = filaImportador.Cells[2].Value.ToString();
                txtboxRFCImportadorDatos.Text = filaImportador.Cells[3].Value.ToString();
                txtboxRepresentanteLegalDatos.Text = filaImportador.Cells[4].Value.ToString();
                txtboxRFCRepresentanteDatos.Text = filaImportador.Cells[5].Value.ToString();
                gboxImportador.Enabled = true;

                btnAltasDatos.Enabled = false;
                btnEliminarDatosImportador.Enabled = false;
                btnGuardarDatosImportador.Enabled = true;

                btnCambiosClic = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Debe selecionar un importador");
                
            }            
        }
        private void btnEliminarDatosImportador_Click(object sender, EventArgs e)
        {
            try
            {                
                bool respuesta;
                Importador miImportador = new Importador()
                {
                    idImportador = int.Parse(filaImportador.Cells[0].Value.ToString())
                };
                DialogResult confirmareliminar = MessageBox.Show("¿Desea eliminar el importador?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirmareliminar == DialogResult.Yes)
                {

                    respuesta = Importador.Instancia.Eliminar(miImportador);
                    if (respuesta) { mostarImportadores(); }
                    limpiarTextBox();
                    filaImportador = new DataGridViewRow();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Debe selecionar un importador");

            }
        }

        private void dgvDatosImportador_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvDatosImportador.Rows[e.RowIndex];
                filaSeleccionada(row);
            }
            
        }
        public void filaSeleccionada(DataGridViewRow row)
        {
            filaImportador = row;
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            gboxImportador.Enabled = false;
            limpiarTextBox();
            btnGuardarDatosImportador.Enabled = false;
            btnCambiosDatos.Enabled = true;
            btnEliminarDatosImportador.Enabled = true;
            btnAltasDatos.Enabled = true;
            btnCambiosClic = false;
        }
        private bool comprobrarVacios()
        {
            List<TextBox> textBoxes = new List<TextBox> {txtboxImportadorDatos, txtboxDomicilioFiscalDatos, txtboxRepresentanteLegalDatos, txtboxRFCImportadorDatos, txtboxRFCRepresentanteDatos};
            bool txtvacio = false;

            foreach (TextBox textBox in textBoxes)
            {
                if (string.IsNullOrWhiteSpace(textBox.Text))
                {
                    txtvacio = true;
                }
            }

            if(txtvacio) { MessageBox.Show("Debe de llenar todos los campos"); }

            return txtvacio;
        }

        private void txtboxRepresentanteLegalDatos_TextChanged(object sender, EventArgs e)
        {

        }

        private void llbDatosImportador_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
    }
}
