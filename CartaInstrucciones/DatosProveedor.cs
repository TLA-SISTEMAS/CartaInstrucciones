using CartaInstrucciones.Modelo;
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
    public partial class DatosProveedor : Form
    {
        LinkLabel linkVisitedLabel = new LinkLabel();
        DataGridViewRow filaProveedor = new DataGridViewRow();
        bool btnCambiosClic = false;
        private static DatosProveedor _instancia;


        public DatosProveedor()
        {
            InitializeComponent();
            linkVisitedLabel = llbDatosProveedor;
        }

        public static DatosProveedor Instancia
        {
            get { 
                if (_instancia == null || _instancia.IsDisposed)
                {
                    _instancia = new DatosProveedor();
                }
                return _instancia; }
        }

        private void DatosProveedor_Load(object sender, EventArgs e)
        {
            
            dgvDatosProveedor.DataSource = null;
            dgvDatosProveedor.DataSource = Proveedor.Instancia.Listar();
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

        private void llbDatosImportador_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            DatosImportador.Instancia.Show();
            //CambiarSeleccion(llbDatosImportador, linkVisitedLabel);
        }


        private void DatosProveedor_FormClosed(object sender, FormClosedEventArgs e)
        {
            Program.cartaInstrucciones.Close();
        }

        private void btnAltas_Click(object sender, EventArgs e)
        {
            gboxDatosProveedor.Enabled = true;
            btnGuardar.Enabled = true;
            btnCambios.Enabled = false;
            btnEliminar.Enabled = false;
            btnCambiosClic = false;
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            bool respuesta;
            if (!btnCambiosClic)
            {
                Proveedor miProveedor = new Proveedor()
                {
                    Nombre = txtboxNombreProveedor.Text.Trim(),
                    Direccion = txtboxDireccion.Text.Trim(),
                    Taxid = txtboxTAXID.Text.Trim()
                };
                if (!comprobrarVacios())
                {
                    btnAltas.Enabled = true;
                    btnGuardar.Enabled = false;
                    btnCambios.Enabled = true;
                    btnEliminar.Enabled = true;
                    gboxDatosProveedor.Enabled = false;

                    respuesta = Proveedor.Instancia.Guardar(miProveedor);
                    if (respuesta) { mostrarProveedores(); }
                    limpiarTextbox();
                }
            }
            else
            {
                Proveedor miProveedor = new Proveedor()
                {
                    ID = int.Parse(filaProveedor.Cells[0].Value.ToString()),
                    Nombre = txtboxNombreProveedor.Text.Trim(),
                    Direccion = txtboxDireccion.Text.Trim(),
                    Taxid = txtboxTAXID.Text.Trim()
                };
                if (!comprobrarVacios())
                {
                    btnAltas.Enabled = true;
                    btnGuardar.Enabled = false;
                    btnCambios.Enabled = true;
                    btnEliminar.Enabled = true;
                    gboxDatosProveedor.Enabled = false;
                    btnCambiosClic = false;
                    respuesta = Proveedor.Instancia.Cambios(miProveedor);
                    if (respuesta) { mostrarProveedores(); }
                    limpiarTextbox();
                }
                filaProveedor = new DataGridViewRow();

            }
        }

        private void btnCambios_Click(object sender, EventArgs e)
        {
            try
            {
                txtboxNombreProveedor.Text = filaProveedor.Cells[1].Value.ToString();
                txtboxDireccion.Text = filaProveedor.Cells[2].Value.ToString();
                txtboxTAXID.Text = filaProveedor.Cells[3].Value.ToString();

                gboxDatosProveedor.Enabled = true;

                btnAltas.Enabled = false;
                btnGuardar.Enabled = true;
                btnEliminar.Enabled = true;

                btnCambiosClic = true;
            }
            catch(Exception)
            {
                MessageBox.Show("Debe selecionar un importador");
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            try
            {
                bool respuesta;
                Proveedor miProveedor = new Proveedor()
                {
                    ID = int.Parse(filaProveedor.Cells[0].Value.ToString()),
                };
                DialogResult confirmarEliminar = MessageBox.Show("¿Desea eliminar el Proveedor?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if(confirmarEliminar == DialogResult.Yes) 
                {
                    respuesta = Proveedor.Instancia.Eliminar(miProveedor);
                    if (respuesta) { mostrarProveedores(); }
                    limpiarTextbox();
                    filaProveedor = new DataGridViewRow();
                }
            }
            catch(Exception)
            {
                MessageBox.Show("Debe selecionar un Proveedor");
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            limpiarTextbox();
            gboxDatosProveedor.Enabled = false;
            btnAltas.Enabled = true;
            btnGuardar.Enabled = false;
            btnCambios.Enabled = true;
            btnEliminar.Enabled = true;
            btnCambiosClic = false;
        }

        private void limpiarTextbox()
        {
            txtboxDireccion.Text = null;
            txtboxNombreProveedor.Text = null;
            txtboxTAXID.Text = null;
        }

        private bool comprobrarVacios()
        {
            List<TextBox> textBoxes = new List<TextBox> { txtboxNombreProveedor, txtboxDireccion, txtboxTAXID};
            bool txtvacio = false;

            foreach (TextBox textBox in textBoxes)
            {
                if (string.IsNullOrWhiteSpace(textBox.Text))
                {
                    txtvacio = true;
                }
            }

            if (txtvacio) { MessageBox.Show("Debe de llenar todos los campos"); }

            return txtvacio;
        }

        public void mostrarProveedores()
        {
            dgvDatosProveedor.DataSource = null;
            dgvDatosProveedor.DataSource = Proveedor.Instancia.Listar();
        }

        private void dgvDatosProveedor_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvDatosProveedor.Rows[e.RowIndex];
                filaSeleccionada(row);
            }
        }
        private void filaSeleccionada(DataGridViewRow row)
        {
            filaProveedor = row;
        }
    }
}
