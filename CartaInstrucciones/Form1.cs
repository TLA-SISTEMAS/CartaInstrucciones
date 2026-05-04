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
            linkVisitedLabel = llbCartaInstrucciones;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            linkVisitedLabel = llbCartaInstrucciones;
            patenteListar();
            importadorListar();
        }

        private void llbDatosImportador_MouseEnter(object sender, EventArgs e)
        {
            if (!llbDatosImportador.LinkVisited) llbDatosImportador.LinkColor = Color.Blue;
        }

        private void llbDatosImportador_MouseLeave(object sender, EventArgs e)
        {
            llbDatosImportador.LinkColor = Color.Black;
        }

        private void llbCartaInstrucciones_MouseEnter(object sender, EventArgs e)
        {
            if (!llbCartaInstrucciones.LinkVisited) llbCartaInstrucciones.LinkColor = Color.Blue;
        }

        private void llbCartaInstrucciones_MouseLeave(object sender, EventArgs e)
        {
            llbCartaInstrucciones.LinkColor = Color.Black;
        }

        private void llbDatosProveedor_MouseEnter(object sender, EventArgs e)
        {
            if (!llbDatosProveedor.LinkVisited) llbDatosProveedor.LinkColor = Color.Blue;
        }

        private void llbDatosProveedor_MouseLeave(object sender, EventArgs e)
        {
            llbDatosProveedor.LinkColor = Color.Black;
        }


        private void llbDatosImportador_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            DatosImportador.Instancia.Show();
        }

        private void llbDatosProveedor_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            DatosProveedor.Instancia.Show();
        }

        
        public void patenteListar()
        {
            foreach (var item in Patente.Instancia.ListarPatente())
            {
                cboxPatente.Items.Add(item.idPatente);
            }
        }

        private void cboxPatente_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboxAduana.Items.Clear();
            cboxAduana.Text = null;
            string patenteSelecionada = cboxPatente.Text;
            foreach (var item in Aduana.Instancia.ListarAduana(patenteSelecionada))
            {
                cboxAduana.Items.Add(item.IDAduana);
            }
        }

        public void importadorListar()
        {
            string idImpo, nombreImpo, idnomImpo;
            foreach (var item in Importador.Instancia.Listar())
            {
                
                idImpo = item.idImportador.ToString();
                nombreImpo = item.nombreImportador.ToString();

                idnomImpo = idImpo + " | " + nombreImpo;
                cboxImportador.Items.Add(idnomImpo);
            }
        }
        private void cboxImportador_SelectedIndexChanged(object sender, EventArgs e)
        {

            int indice = cboxImportador.Text.IndexOf("|");
            int idImportador = int.Parse(cboxImportador.Text.Substring(0, indice - 1));

            //MessageBox.Show(idImportador.ToString());
            foreach (var item in Importador.Instancia.ListarunImportador(idImportador))
            {
                txtboxDomicilioFiscaCartaInstrucciones.Text = item.domicilio.ToString();
                txtboxRFCImportadorCarta.Text = item.rfc.ToString();
                txtboxRepresentanteLegalCarta.Text = item.representanteLegal.ToString();
                txtboxRFCRepresentanteCarta.Text = item.rfcRepresentante.ToString();
            }
            gboxGenerarlesEmbarque.Enabled = true;
        }

        private void cboxImportador_MouseClick(object sender, MouseEventArgs e)
        {
            cboxImportador.Items.Clear();
            importadorListar();
        }

        private void btnCapturarFacturas_Click(object sender, EventArgs e)
        {
            cantidadFacturasCapturarGlobal = int.Parse(nudFacturas.Value.ToString());
            
            if (nudFacturas.Value <= 10)
            {
                if (dgvFacturasIndividual.Rows.Count <= cantidadFacturasCapturarGlobal) 
                {
                    gboxFacturaIndividual.Enabled = true;
                    dgvFacturasIndividual.Visible = true;
                    dgvFacturasLotes.Visible = false;
                    gboxFacturaLote.Enabled = false;
                    lblFacturasCapturardasIndividual.Visible = true;
                    lblFacturasCapturadasLotes.Visible = false;
                    cboxProveedor1.Items.AddRange(proveedorListar().ToArray());
                    lblFacturasaCapturar.Text = cantidadFacturasCapturarGlobal.ToString();
                } 
                else
                {
                    MessageBox.Show("El numero de facturas por capturar no puede ser menor al numero ya capturado");
                }
            }
            else
            {
                if (int.Parse(lblFacturasCapturadasLotes.Text) <= cantidadFacturasCapturarGlobal)
                {
                    gboxFacturaLote.Enabled = true;
                    gboxFacturaIndividual.Enabled = false;
                    dgvFacturasIndividual.Visible = false;
                    dgvFacturasLotes.Visible = true;
                    lblFacturasCapturardasIndividual.Visible = false;
                    lblFacturasCapturadasLotes.Visible = true;
                    cboxProveedor2.Items.AddRange(proveedorListar().ToArray());
                    lblFacturasaCapturar.Text = cantidadFacturasCapturarGlobal.ToString();
                }
                else
                {
                    MessageBox.Show("El numero de facturas por capturar no puede ser menor al numero ya capturado");
                }
            }
        }

        private List<string> proveedorListar()
        {
            List<string> listProveedores = new List<string>();
            string idProve, nombreProve, idnomProve;
            foreach (var item in Proveedor.Instancia.Listar())
            {
                idProve = item.ID.ToString();
                nombreProve = item.Nombre.ToString();

                idnomProve = idProve + " | " + nombreProve;
                //listProveedores.Add(idnomProve);
                listProveedores.Add(idnomProve);
            }

            return listProveedores;
        }


        private void btnGuardar_Click(object sender, EventArgs e)
        {
            List<TextBox> list = new List<TextBox> { txtboxFactura, txtboxValorFactura1, txtboxDescripcionGenericaIndividual };
            if (!comprobarVacios(list, cboxProveedor1))
            {
                int indice = cboxProveedor1.Text.IndexOf("|");
                int idProveedor = int.Parse(cboxProveedor1.Text.Substring(0, indice - 1));
                string nombreProveedor = cboxProveedor1.Text.Substring(indice + 2);
                if (dgvFacturasIndividual.Enabled)
                {
                    if (int.Parse(lblFacturasCapturardasIndividual.Text) < nudFacturas.Value)
                    {
                        dgvFacturasIndividual.Rows.Add(txtboxFactura.Text, txtboxValorFactura1.Text, nombreProveedor, txtboxDescripcionGenericaIndividual.Text, idProveedor.ToString());
                        txtboxFactura.Text = null;
                        txtboxValorFactura1.Text = null;
                        txtboxDescripcionGenericaIndividual.Text= null;
                        dgvFacturasIndividual.Enabled = true;
                        btnEliminar.Enabled = true;

                        lblFacturasCapturardasIndividual.Text = (int.Parse(lblFacturasCapturardasIndividual.Text) + 1).ToString();
                    }
                    else
                    {
                        MessageBox.Show("Se han capturado todas las Facturas");
                    }
                    
                }
                else
                {
                    dgvFacturasIndividual.CurrentRow.Cells[0].Value = txtboxFactura.Text;
                    dgvFacturasIndividual.CurrentRow.Cells[1].Value = txtboxValorFactura1.Text;
                    dgvFacturasIndividual.CurrentRow.Cells[2].Value = nombreProveedor;
                    dgvFacturasIndividual.CurrentRow.Cells[3].Value = txtboxDescripcionGenericaIndividual.Text;
                    dgvFacturasIndividual.CurrentRow.Cells[4].Value = idProveedor.ToString();
                    txtboxFactura.Text = null;
                    txtboxValorFactura1.Text = null;
                    txtboxDescripcionGenericaIndividual.Text = null;
                    dgvFacturasIndividual.Enabled = true;
                    btnEliminar.Enabled = true;
                }

            }   
        }

        private bool comprobarVacios(List<TextBox> textBoxes, ComboBox cboxproveedor)
        {
            bool respuesta = false;

            foreach (TextBox textbox in textBoxes)
            {
                if (string.IsNullOrWhiteSpace(textbox.Text))
                {
                    respuesta = true;
                }
            }
            if(string.IsNullOrEmpty(cboxproveedor.Text))
            {
                respuesta = true;
            }
            if (respuesta) { MessageBox.Show("Debe de llenar todos los campos"); }
            return respuesta;
        }

        private void btnCambios_Click(object sender, EventArgs e)
        {
            try {
                if (dgvFacturasIndividual.CurrentRow != null && !dgvFacturasIndividual.CurrentRow.IsNewRow)
                {
                    string idProveedor = dgvFacturasIndividual.CurrentRow.Cells[4].Value.ToString();
                    string nombreProvedor = dgvFacturasIndividual.CurrentRow.Cells[2].Value.ToString();
                    //MessageBox.Show(idProveedor + " | " + nombreProvedor);
                    txtboxFactura.Text = dgvFacturasIndividual.CurrentRow.Cells[0].Value.ToString();
                    txtboxValorFactura1.Text = dgvFacturasIndividual.CurrentRow.Cells[1].Value.ToString();
                    cboxProveedor1.Text = idProveedor + " | " + nombreProvedor;
                    txtboxDescripcionGenericaIndividual.Text = dgvFacturasIndividual.CurrentRow.Cells[3].Value.ToString();
                    dgvFacturasIndividual.Enabled = false;
                    btnEliminar.Enabled = false;
                }
            }
            catch (Exception)
            {

            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (dgvFacturasIndividual.CurrentRow != null && !dgvFacturasIndividual.CurrentRow.IsNewRow)
            {
                dgvFacturasIndividual.Rows.Remove(dgvFacturasIndividual.CurrentRow);
                lblFacturasCapturardasIndividual.Text = (int.Parse(lblFacturasCapturardasIndividual.Text) - 1).ToString();
            }
        }

        private void cboxProveedor1_Click(object sender, EventArgs e)
        {
            cboxProveedor1.Items.Clear();
            cboxProveedor1.Items.AddRange(proveedorListar().ToArray());
        }

        private void btnGuardarLotes_Click(object sender, EventArgs e)
        {
            List<TextBox> list = new List<TextBox> { txtboxValorFactura2, txtboxDescripcionGenericaILotes};
            if (!comprobarVacios(list, cboxProveedor2))
            {
                int indice = cboxProveedor2.Text.IndexOf("|");
                int idProveedor = int.Parse(cboxProveedor2.Text.Substring(0, indice - 1));
                string nombreProveedor = cboxProveedor2.Text.Substring(indice + 2);
                if (dgvFacturasLotes.Enabled)
                {
                    int facturasTotales = int.Parse(lblFacturasCapturadasLotes.Text) + int.Parse(nudCantidadFacturasProveedor.Value.ToString());
                    if (nudCantidadFacturasProveedor.Value <= int.Parse(lblFacturasaCapturar.Text) && facturasTotales <= int.Parse(lblFacturasaCapturar.Text))
                    {
                        dgvFacturasLotes.Rows.Add(nombreProveedor, nudCantidadFacturasProveedor.Value.ToString(), txtboxValorFactura2.Text, txtboxDescripcionGenericaILotes.Text, idProveedor);
                        cboxProveedor2.Text = null;                        
                        txtboxValorFactura2.Text = null;
                        txtboxDescripcionGenericaILotes.Text = null;
                        dgvFacturasLotes.Enabled = true;
                        btnEliminarLotes.Enabled = true;

                        lblFacturasCapturadasLotes.Text = ( int.Parse(lblFacturasCapturadasLotes.Text) + int.Parse(nudCantidadFacturasProveedor.Value.ToString())).ToString();

                        nudCantidadFacturasProveedor.Value = 1;
                    }
                    else
                    {
                        MessageBox.Show("La cantidad de facturas no puede ser mayor a \"Facturas a Capturar\"");
                    }
                }
                else
                {
                    int cantidadFacturasActual = int.Parse(dgvFacturasLotes.CurrentRow.Cells[1].Value.ToString());
                    int cantidadFacturasNuevo = int.Parse(nudCantidadFacturasProveedor.Value.ToString());
                    if(cantidadFacturasNuevo <= cantidadFacturasActual)
                    {
                        int cantidadFacturasRestadas = cantidadFacturasActual - cantidadFacturasNuevo;
                        lblFacturasCapturadasLotes.Text = (int.Parse(lblFacturasCapturadasLotes.Text) - cantidadFacturasRestadas).ToString();
                    }
                    else
                    {
                        int cantidadFacturasSumadas = cantidadFacturasNuevo - cantidadFacturasActual;
                        lblFacturasCapturadasLotes.Text = (int.Parse(lblFacturasCapturadasLotes.Text) + cantidadFacturasSumadas).ToString();
                    }

                    dgvFacturasLotes.CurrentRow.Cells[0].Value = nombreProveedor;
                    dgvFacturasLotes.CurrentRow.Cells[1].Value = nudCantidadFacturasProveedor.Value.ToString();
                    dgvFacturasLotes.CurrentRow.Cells[2].Value = txtboxValorFactura2.Text;
                    dgvFacturasLotes.CurrentRow.Cells[3].Value = txtboxDescripcionGenericaILotes.Text;
                    dgvFacturasLotes.CurrentRow.Cells[4].Value = idProveedor.ToString();
                    cboxProveedor2.Text = null;
                    nudCantidadFacturasProveedor.Value = 1;
                    txtboxValorFactura2.Text = null;
                    txtboxDescripcionGenericaILotes.Text = null;
                    dgvFacturasLotes.Enabled = true;
                    btnEliminarLotes.Enabled = true;
                }

            }
        }

        private void btnCabiosLotes_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvFacturasLotes.CurrentRow != null && !dgvFacturasLotes.CurrentRow.IsNewRow)
                {
                    string idProveedor = dgvFacturasLotes.CurrentRow.Cells[4].Value.ToString();
                    string nombreProvedor = dgvFacturasLotes.CurrentRow.Cells[0].Value.ToString();
                    //MessageBox.Show(idProveedor + " | " + nombreProvedor);
                    cboxProveedor2.Text = idProveedor + " | " + nombreProvedor;
                    nudCantidadFacturasProveedor.Value = int.Parse(dgvFacturasLotes.CurrentRow.Cells[1].Value.ToString());
                    txtboxValorFactura2.Text = dgvFacturasLotes.CurrentRow.Cells[2].Value.ToString();
                    txtboxDescripcionGenericaILotes.Text = dgvFacturasLotes.CurrentRow.Cells[3].Value.ToString();
                    dgvFacturasLotes.Enabled = false;
                    btnEliminarLotes.Enabled = false;
                }
            }
            catch (Exception)
            {

            }
        }

        private void btnEliminarLotes_Click(object sender, EventArgs e)
        {
            if (dgvFacturasLotes.CurrentRow != null && !dgvFacturasLotes.CurrentRow.IsNewRow)
            {
                int cantidadFacturasEliminadas = int.Parse(dgvFacturasLotes.CurrentRow.Cells[1].Value.ToString());
                lblFacturasCapturadasLotes.Text = (int.Parse(lblFacturasCapturadasLotes.Text) - cantidadFacturasEliminadas).ToString();
                dgvFacturasLotes.Rows.Remove(dgvFacturasLotes.CurrentRow);
            }
        }

        private void cboxProveedor2_Click(object sender, EventArgs e)
        {
            cboxProveedor2.Items.Clear();
            cboxProveedor2.Items.AddRange(proveedorListar().ToArray());
        }


        private void txtboxValorFactura1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxValorFactura2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxIncreFlete_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxIncreSeguro_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxIncreEmbalaje_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxIncreOtrosgastos_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxDreceTransporte_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxDreceSeguro_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxDreceCarga_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxDreceEmbalaje_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private void txtboxDreceOtrosGastos_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Ignora la tecla
            }
        }

        private TextBox formatoMoneda (TextBox textBox)
        {

            if (decimal.TryParse(textBox.Text, out decimal valor))
            {
                // Buscamos cuántos decimales escribió el usuario realmente
                string[] partes = textBox.Text.Split(System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]);

                if (partes.Length > 1)
                {
                    int cantidadDecimales = partes[1].Length;
                    // Formateamos con el número exacto de decimales que ingresó
                    textBox.Text = valor.ToString("N" + cantidadDecimales);
                }
                else
                {
                    // Si no puso decimales, podemos mostrar "N0" o simplemente "N"
                    textBox.Text = valor.ToString("N0");
                }
            }

            return textBox;
        }

        private void txtboxValorFactura_Leave(object sender, EventArgs e)
        {
            txtboxValorFactura1 = formatoMoneda(txtboxValorFactura1);
        }

        private void txtboxValorFactura2_Leave(object sender, EventArgs e)
        {
            txtboxValorFactura2 = formatoMoneda(txtboxValorFactura2);
        }

        private void txtboxIncreFlete_Leave(object sender, EventArgs e)
        {
            txtboxIncreFlete = formatoMoneda(txtboxIncreFlete);
        }

        private void txtboxIncreSeguro_Leave(object sender, EventArgs e)
        {
            txtboxIncreSeguro = formatoMoneda(txtboxIncreSeguro);
        }

        private void txtboxIncreEmbalaje_Leave(object sender, EventArgs e)
        {
            txtboxIncreEmbalaje = formatoMoneda(txtboxIncreEmbalaje);
        }

        private void txtboxIncreOtrosgastos_Leave(object sender, EventArgs e)
        {
            txtboxIncreOtrosgastos = formatoMoneda(txtboxIncreOtrosgastos);
        }

        private void txtboxDreceTransporte_Leave(object sender, EventArgs e)
        {
            txtboxDreceTransporte = formatoMoneda(txtboxDreceTransporte);
        }

        private void txtboxDreceSeguro_Leave(object sender, EventArgs e)
        {
            txtboxDreceSeguro = formatoMoneda(txtboxDreceSeguro);
        }

        private void txtboxDreceCarga_Leave(object sender, EventArgs e)
        {
            txtboxDreceCarga = formatoMoneda(txtboxDreceCarga);
        }

        private void txtboxDreceEmbalaje_Leave(object sender, EventArgs e)
        {
            txtboxDreceEmbalaje = formatoMoneda(txtboxDreceEmbalaje);
        }

        private void txtboxDreceOtrosGastos_Leave(object sender, EventArgs e)
        {
            txtboxDreceOtrosGastos = formatoMoneda(txtboxDreceOtrosGastos);
        }

        private void btnGenerarCarta_Click(object sender, EventArgs e)
        {
            SaveFileDialog rutaCarta = new SaveFileDialog();

            rutaCarta.Filter = "Documneto de Word |*.docx";
            rutaCarta.Title = "Guardar Carta de Instrucciones";

            rutaCarta.ShowDialog();
            string ruta;
            if (rutaCarta.FileName != "")
            {
                //Iniciar Aplicacion
                objWord.Application objAplication = new objWord.Application();
                //Crear un Nuevo documento
                objWord.Document objDocument = objAplication.Documents.Add();

                objWord.PageSetup pageSetup = objDocument.PageSetup;

                float cmToPoints = 28.35f;
                pageSetup.TopMargin = 1.59f * cmToPoints;
                pageSetup.BottomMargin = 2.22f * cmToPoints;
                pageSetup.LeftMargin = 1f * cmToPoints;
                pageSetup.RightMargin = 1f * cmToPoints;


                objWord.Paragraph objParrafo1 = objDocument.Content.Paragraphs.Add(Type.Missing);

                //TAMAÑO DE LETRA
                objParrafo1.Range.Font.Size = 8;
                //AGREGAR TEXTO A UN PARRAFO
                objParrafo1.Range.Text = "CARTA DE INSTRUCCIONES";
                //TIPO DE LETRA
                objParrafo1.Range.Font.Name = "Times New Roman";
                //CENTRAR UN PARRAFO
                objParrafo1.Alignment = objWord.WdParagraphAlignment.wdAlignParagraphCenter;

                objParrafo1.Range.InsertParagraphAfter();

                //AA

                objParrafo1.Range.Text = "A.A " + txtboxAA.Text;
                objParrafo1.Range.Font.Name = "Arial";
                objParrafo1.Range.Font.Size = 8;
                objParrafo1.Alignment = objWord.WdParagraphAlignment.wdAlignParagraphLeft;
                objParrafo1.Range.InsertParagraphAfter();

                //PATENTE ADUANAL

                objParrafo1.Range.Text = "PATENTE ADUANAL " + cboxPatente.Text;
                objParrafo1.Range.Font.Name = "Arial";
                objParrafo1.Range.Font.Size = 8;
                objParrafo1.Range.InsertParagraphAfter();

                //ADUANA DE

                objParrafo1.Range.Text = "ADUANA DE " + cboxAduana.Text;
                objParrafo1.Range.Font.Name = "Arial";
                objParrafo1.Range.Font.Size = 8;
                objParrafo1.Range.InsertParagraphAfter();

                //Fecha


                objParrafo1.Range.Text = DateTime.Now.ToString("dd/MM/yyyy");
                objParrafo1.Range.Font.Name = "Arial";
                objParrafo1.Range.Font.Size = 8;
                objParrafo1.Alignment = objWord.WdParagraphAlignment.wdAlignParagraphRight;
                objParrafo1.Range.InsertParagraphAfter();

                //PRESENTE

                objParrafo1.Range.Text = "P R E S E N T E";
                objParrafo1.Range.Font.Name = "Arial";
                objParrafo1.Range.Font.Size = 8;
                objParrafo1.Alignment = objWord.WdParagraphAlignment.wdAlignParagraphLeft;
                objParrafo1.Range.InsertParagraphAfter();

                //Texto Largo

                objWord.Paragraph textoLargo = objDocument.Paragraphs.Add();
                objWord.Range rangeTextoLargo = textoLargo.Range;

                int indice = cboxImportador.Text.IndexOf("|");
                string nombreImportador = cboxImportador.Text.Substring(indice + 1);
                
                rangeTextoLargo.Font.Size = 8;
                rangeTextoLargo.Font.Name = "Arial";
                
                rangeTextoLargo.Text = txtboxRepresentanteLegalCarta.Text.ToUpper();
                rangeTextoLargo.Font.Bold = 1;
                rangeTextoLargo.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeTextoLargo.Text = " , Representante legal de la Empresa ";
                rangeTextoLargo.Font.Bold = 0;
                rangeTextoLargo.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeTextoLargo.Text = nombreImportador.ToUpper();
                rangeTextoLargo.Font.Bold = 1;
                rangeTextoLargo.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeTextoLargo.Text = " , con domicilio Fiscal ubicado en domicilio ";
                rangeTextoLargo.Font.Bold = 0;
                rangeTextoLargo.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeTextoLargo.Text = txtboxDomicilioFiscaCartaInstrucciones.Text.ToUpper();
                rangeTextoLargo.Font.Bold = 1;
                rangeTextoLargo.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeTextoLargo.Text = " y  R.F.C. ";
                rangeTextoLargo.Font.Bold = 0;
                rangeTextoLargo.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeTextoLargo.Text = txtboxRFCImportadorCarta.Text.ToUpper();
                rangeTextoLargo.Font.Bold = 1;
                rangeTextoLargo.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeTextoLargo.Text = " con base en lo prescrito en los Arts. 40,  41  y demás correlativos de la ley Aduanera, El Suscrito por medio de la presente Carta de  Encomienda, le confiero Poder amplio, cumplido y bastante, para que en mi Nombre y con la calidad de Representante, realice  todos los trámites relativos al despacho de mercancía ante La Aduana correspondiente a la siguiente Información:";
                rangeTextoLargo.Font.Bold = 0;
                rangeTextoLargo.InsertParagraphAfter();

                objWord.Paragraph objParrafoTitulos = objDocument.Content.Paragraphs.Add(Type.Missing);

                objParrafoTitulos.Range.Font.Size = 9;
                objParrafoTitulos.Range.Text = "GENERALES DE EMBARQUE";
                objParrafoTitulos.Range.Bold = 1;
                //Modificar espaciado entre parrafos
                objParrafoTitulos.Format.SpaceAfter = 8f;
                objParrafoTitulos.Range.InsertParagraphAfter();

                object missing = System.Reflection.Missing.Value;

                objWord.Range rangetablaFacturasGeneralEmbarque = objDocument.Paragraphs.Add(ref missing).Range;
                
                objWord.Cell celda;
                List<float> tamanios = new List<float>();
                List<string> listaidProvedores = new List<string>();

                //TALBA FACTURAS
                objWord.Paragraph parrafoEnter = objDocument.Paragraphs.Add(ref missing);
                
                if (dgvFacturasIndividual.Enabled)
                {                
                    int filas = dgvFacturasIndividual.Rows.Count + 1;
                    //int filas = 5;
                    int columnas = 4;
                    objWord.Table tablaFacturasGeneralEmbarque = objDocument.Tables.Add(rangetablaFacturasGeneralEmbarque, filas, columnas);

                    tamanios.Clear();
                    tamanios.AddRange(new float[] { 1.99f, 2f, 3.75f, 11.75f });

                    tablaFacturasGeneralEmbarque = asignarTamanioColumna(tablaFacturasGeneralEmbarque, columnas, tamanios);


                    tablaFacturasGeneralEmbarque.Rows[1].Height = 0.77f * cmToPoints;
                    tablaFacturasGeneralEmbarque.Borders.Enable = 1;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 1);
                    celda = formatoCelda(celda, "Factura", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 1).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 2);
                    celda = formatoCelda(celda, "Valor Factura", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 2).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 3);
                    celda = formatoCelda(celda, "Proveedor", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 3).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 3).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 4);
                    celda = formatoCelda(celda, "Descripción Genérica de la Mercancía", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 4).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 4).Range.ParagraphFormat.SpaceAfter = 0;

                    for (int r = 2; r <= filas; r++)
                    {
                        for (int c = 1; c <= columnas; c++)
                        {
                            tablaFacturasGeneralEmbarque.Cell(r, c).Range.Text = dgvFacturasIndividual.Rows[r - 2].Cells[c - 1].Value.ToString();
                            tablaFacturasGeneralEmbarque.Cell(r, c).Range.Bold = 0;

                            //tablaGeneralEmbarque.Cell(r, c).Range.Text = $"F{r-2}C{c-1}";
                        }
                       // MessageBox.Show("Fuera del IF "+dgvFacturasIndividual.Rows[r - 2].Cells[4].Value.ToString());
                        if (!listaidProvedores.Contains(dgvFacturasIndividual.Rows[r - 2].Cells[4].Value.ToString()))
                        {
                          //  MessageBox.Show("Dentro del IF " + dgvFacturasIndividual.Rows[r - 2].Cells[4].Value.ToString());
                            listaidProvedores.Add(dgvFacturasIndividual.Rows[r - 2].Cells[4].Value.ToString());
                        }
                    }

                    parrafoEnter.Range.Text = "";
                    parrafoEnter.Format.SpaceAfter = 0f;
                    parrafoEnter.Range.InsertParagraphAfter();

                    objParrafoTitulos.Range.Font.Size = 8;
                    objParrafoTitulos.Range.Text = "Proveedores";
                    objParrafoTitulos.Range.Bold = 1;
                    //Modificar espaciado entre parrafos
                    objParrafoTitulos.Format.SpaceAfter = 0f;
                    objParrafoTitulos.Format.SpaceBefore = 0f;
                    objParrafoTitulos.Range.InsertParagraphAfter();

                    objWord.Range rangetablaProveedores = objDocument.Paragraphs.Add(ref missing).Range;
                    objWord.Table tablaProveedores = objDocument.Tables.Add(rangetablaProveedores, listaidProvedores.Count + 1, 3);

                    tablaProveedores.Borders.Enable = 1;

                    celda = tablaProveedores.Cell(1, 1);
                    celda = formatoCelda(celda, "Nombre", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaProveedores.Cell(1, 1).BottomPadding = 0;
                    tablaProveedores.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaProveedores.Cell(1, 2);
                    celda = formatoCelda(celda, "Direccion", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaProveedores.Cell(1, 1).BottomPadding = 0;
                    tablaProveedores.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaProveedores.Cell(1, 3);
                    celda = formatoCelda(celda, "TAXID", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaProveedores.Cell(1, 1).BottomPadding = 0;
                    tablaProveedores.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    int renglonProveedor = 2;
                    foreach(var item in listaidProvedores)
                    {
                        //MessageBox.Show("Item de la lista Id Proveedores" + item.ToString() +"Posiconado en " + listaidProvedores.IndexOf(item) + " Colocado en el reglon " + renglonProveedor );
                        Proveedor miProveedor = Proveedor.Instancia.ListarunProvedor(int.Parse(item));

                        //MessageBox.Show(miProveedor.Nombre +" "+ miProveedor.Direccion +" "+ miProveedor.Taxid);

                        tablaProveedores.Cell(renglonProveedor, 1).Range.Text = miProveedor.Nombre;
                        tablaProveedores.Cell(renglonProveedor, 2).Range.Text = miProveedor.Direccion;
                        tablaProveedores.Cell(renglonProveedor, 3).Range.Text = miProveedor.Taxid;

                        //MessageBox.Show("Rengo "+renglonProveedor+" "+tablaProveedores.Cell(renglonProveedor, 1).Range.Text +" "+ tablaProveedores.Cell(renglonProveedor, 2).Range.Text +" "+ tablaProveedores.Cell(renglonProveedor, 3).Range.Text);
                        renglonProveedor++;

                    }

                }
                if (dgvFacturasLotes.Enabled)
                {
                    int filas = dgvFacturasLotes.Rows.Count + 1;
                    int columnas = 4;

                    objWord.Table tablaFacturasGeneralEmbarque = objDocument.Tables.Add(rangetablaFacturasGeneralEmbarque, filas, columnas);
                    tamanios.Clear();
                    tamanios.AddRange(new float[] { 3.99f, 1.5f, 3f, 11f});

                    tablaFacturasGeneralEmbarque = asignarTamanioColumna(tablaFacturasGeneralEmbarque, columnas, tamanios);

                    tablaFacturasGeneralEmbarque.Rows[1].Height = 0.77f * cmToPoints;
                    tablaFacturasGeneralEmbarque.Borders.Enable = 1;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 1);
                    celda = formatoCelda(celda, "Proveedor", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 1).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 2);
                    celda = formatoCelda(celda, "Cantidad Facturas", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 2).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 3);
                    celda = formatoCelda(celda, "Valor de las Facturas", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 3).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 3).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaFacturasGeneralEmbarque.Cell(1, 4);
                    celda = formatoCelda(celda, "Descripción Genérica de la Mercancía", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaFacturasGeneralEmbarque.Cell(1, 4).BottomPadding = 0;
                    tablaFacturasGeneralEmbarque.Cell(1, 4).Range.ParagraphFormat.SpaceAfter = 0;

                    for (int r = 2; r <= filas; r++)
                    {
                        for (int c = 1; c <= columnas; c++)
                        {
                            tablaFacturasGeneralEmbarque.Cell(r, c).Range.Text = dgvFacturasLotes.Rows[r - 2].Cells[c - 1].Value.ToString();
                            tablaFacturasGeneralEmbarque.Cell(r, c).Range.Bold = 0;
                            //tablaGeneralEmbarque.Cell(r, c).Range.Text = $"F{r-2}C{c-1}";
                        }
                        //MessageBox.Show("Fuera del IF " + dgvFacturasIndividual.Rows[r - 2].Cells[4].Value.ToString());
                        if (!listaidProvedores.Contains(dgvFacturasLotes.Rows[r - 2].Cells[4].Value.ToString()))
                        {
                           // MessageBox.Show("Dentro del IF " + dgvFacturasIndividual.Rows[r - 2].Cells[4].Value.ToString());
                            listaidProvedores.Add(dgvFacturasLotes.Rows[r - 2].Cells[4].Value.ToString());
                        }
                    }

                    parrafoEnter.Range.Text = "";
                    parrafoEnter.Format.SpaceAfter = 0f;
                    parrafoEnter.Range.InsertParagraphAfter();

                    objParrafoTitulos.Range.Font.Size = 8;
                    objParrafoTitulos.Range.Text = "Proveedores";
                    objParrafoTitulos.Range.Bold = 1;
                    //Modificar espaciado entre parrafos
                    objParrafoTitulos.Format.SpaceAfter = 0f;
                    objParrafoTitulos.Format.SpaceBefore = 0f;
                    objParrafoTitulos.Range.InsertParagraphAfter();

                    objWord.Range rangetablaProveedores = objDocument.Paragraphs.Add(ref missing).Range;
                    objWord.Table tablaProveedores = objDocument.Tables.Add(rangetablaProveedores, listaidProvedores.Count + 1, 3);

                    tablaProveedores.Borders.Enable = 1;

                    celda = tablaProveedores.Cell(1, 1);
                    celda = formatoCelda(celda, "Nombre", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaProveedores.Cell(1, 1).BottomPadding = 0;
                    tablaProveedores.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaProveedores.Cell(1, 2);
                    celda = formatoCelda(celda, "Direccion", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaProveedores.Cell(1, 1).BottomPadding = 0;
                    tablaProveedores.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    celda = tablaProveedores.Cell(1, 3);
                    celda = formatoCelda(celda, "TAXID", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                    tablaProveedores.Cell(1, 1).BottomPadding = 0;
                    tablaProveedores.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;

                    int renglonProveedor = 2;
                    foreach (var item in listaidProvedores)
                    {
                        //MessageBox.Show("Item de la lista Id Proveedores" + item.ToString() +"Posiconado en " + listaidProvedores.IndexOf(item) + " Colocado en el reglon " + renglonProveedor );
                        Proveedor miProveedor = Proveedor.Instancia.ListarunProvedor(int.Parse(item));

                        //MessageBox.Show(miProveedor.Nombre +" "+ miProveedor.Direccion +" "+ miProveedor.Taxid);

                        tablaProveedores.Cell(renglonProveedor, 1).Range.Text = miProveedor.Nombre;
                        tablaProveedores.Cell(renglonProveedor, 2).Range.Text = miProveedor.Direccion;
                        tablaProveedores.Cell(renglonProveedor, 3).Range.Text = miProveedor.Taxid;

                        //MessageBox.Show("Rengo "+renglonProveedor+" "+tablaProveedores.Cell(renglonProveedor, 1).Range.Text +" "+ tablaProveedores.Cell(renglonProveedor, 2).Range.Text +" "+ tablaProveedores.Cell(renglonProveedor, 3).Range.Text);
                        renglonProveedor++;

                    }
                }
                

                //AÑADIR ESPACIO ENTRE TABLAS

                
                parrafoEnter.Range.Text = "";
                parrafoEnter.Format.SpaceAfter = 0f;
                parrafoEnter.Range.InsertParagraphAfter();

                objWord.Range rangetablaGeneralEmbarque = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablaGeneralEmbarque = objDocument.Tables.Add(rangetablaGeneralEmbarque, 4, 2);

                //TABLA GUIA Maritima

                tamanios.Clear();
                tamanios.AddRange(new float[] { 2.49f, 17f} );
                
                tablaGeneralEmbarque = asignarTamanioColumna(tablaGeneralEmbarque, 2, tamanios);

                tablaGeneralEmbarque.Rows[4].Height = 1.09f * cmToPoints;
                tablaGeneralEmbarque.Borders.Enable = 1;

                celda = tablaGeneralEmbarque.Cell(1, 1);
                celda = formatoCelda(celda, "Guía Marítima", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaGeneralEmbarque.Cell(1, 2).Range.Text = txtboxGuiaMaritima.Text.ToUpper();
                tablaGeneralEmbarque.Cell(1, 2).Range.Bold = 0;

                celda = tablaGeneralEmbarque.Cell(2, 1);
                celda = formatoCelda(celda, "Transportista", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaGeneralEmbarque.Cell(2,2).Range.Text = txtboxTransportista.Text.ToUpper();
                tablaGeneralEmbarque.Cell(2, 2).Range.Bold = 0;


                celda = tablaGeneralEmbarque.Cell(3, 1);
                celda = formatoCelda(celda, "Lista de Empaque", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaGeneralEmbarque.Cell(3, 2).Range.Text = txtboxListaEmpaque.Text.ToUpper();
                tablaGeneralEmbarque.Cell(3, 2).Range.Bold = 0;

                celda = tablaGeneralEmbarque.Cell(4, 1);
                celda = formatoCelda(celda, "Observaciones Generales del Embarque", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaGeneralEmbarque.Cell(4, 2).Range.Text = txtboxObservacionesGeneralesEmbarque.Text;
                tablaGeneralEmbarque.Cell(4, 2).Range.Bold = 0;


                objParrafoTitulos.Range.Text = "GENERALES DE DESPACHO ADUANAL";
                objParrafoTitulos.Range.Font.Size = 9;
                objParrafoTitulos.Format.SpaceBefore = 6f;
                objParrafoTitulos.Format.SpaceAfter = 0f;
                objParrafoTitulos.Range.InsertParagraphAfter();

                objParrafoTitulos.Format.SpaceBefore = 0f;

                //TABLA TIPO OPERACION REGIMEN E INCOTERM (TORAI)

                objWord.Range rangetablaGeneralDespachoTORA = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablaGeneralDespachoTORAI = objDocument.Tables.Add(rangetablaGeneralDespachoTORA, 1, 6);

                tablaGeneralDespachoTORAI.Borders.Enable = 1;

                tamanios.Clear();
                tamanios.AddRange(new float[] { 1.78f, 2.46f, 1.57f, 11.07f, 1.57f, 1.04f });
                tablaGeneralDespachoTORAI = asignarTamanioColumna(tablaGeneralDespachoTORAI, 6, tamanios);

                celda = tablaGeneralDespachoTORAI.Cell(1, 1);
                celda = formatoCelda(celda, "Tipo de Operación", 8 ,1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaGeneralDespachoTORAI.Cell(1, 2);

                if(!ckboxR1.Checked)
                {
                    celda = formatoCelda(celda, cboxTipoOperacion.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                }
                else
                {
                    celda = formatoCelda(celda, cboxTipoOperacion.Text + " R1", 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);
                }

                celda = tablaGeneralDespachoTORAI.Cell(1, 3);
                celda = formatoCelda(celda, "Régimen Aduanal", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaGeneralDespachoTORAI.Cell(1, 4);
                celda = formatoCelda(celda, txtboxRegimenAduanal.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphLeft);

                celda = tablaGeneralDespachoTORAI.Cell(1, 5);
                celda = formatoCelda(celda, "Incoterm", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaGeneralDespachoTORAI.Cell(1, 6);
                celda = formatoCelda(celda, cboxIncoterm.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                parrafoEnter.Range.Text = "";
                parrafoEnter.Format.SpaceAfter = 0f;
                parrafoEnter.Range.InsertParagraphAfter();

                objParrafoTitulos.Range.Font.Size = 8;
                objParrafoTitulos.Range.Text = "Incrementables";
                objParrafoTitulos.Range.Bold = 1;
                //Modificar espaciado entre parrafos
                objParrafoTitulos.Format.SpaceAfter = 0f;
                objParrafoTitulos.Format.SpaceBefore = 0f;
                objParrafoTitulos.Range.InsertParagraphAfter();


                //TABLA INCREMENTABLES
                objWord.Range rangetablaIncrementables = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablaIncrementables = objDocument.Tables.Add(rangetablaIncrementables, 1, 8);

                tablaIncrementables.Borders.Enable = 1;

                tamanios.Clear();
                tamanios.AddRange(new float[] { 1.64f, 3f, 1.64f, 3f, 1.64f, 3f, 1.64f, 3f});
                tablaIncrementables = asignarTamanioColumna(tablaIncrementables, 6, tamanios);

                celda = tablaIncrementables.Cell(1, 1);
                celda = formatoCelda(celda, "Flete", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaIncrementables.Cell(1, 2);
                celda = formatoCelda(celda, txtboxIncreFlete.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaIncrementables.Cell(1, 3);
                celda = formatoCelda(celda, "Seguro", 8,1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaIncrementables.Cell(1, 4);
                celda = formatoCelda(celda, txtboxIncreSeguro.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaIncrementables.Cell(1, 5);
                celda = formatoCelda(celda, "Embalaje", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaIncrementables.Cell(1, 6);
                celda = formatoCelda(celda, txtboxIncreEmbalaje.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaIncrementables.Cell(1, 7);
                celda = formatoCelda(celda, "Otros", 8, 0, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaIncrementables.Cell(1, 8);
                celda = formatoCelda(celda, txtboxIncreOtrosgastos.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                objWord.Paragraph fechaIncrementables = objDocument.Paragraphs.Add();
                objWord.Range rangeFechaIncrementables = fechaIncrementables.Range;

                rangeFechaIncrementables.Font.Size = 8;
                rangeFechaIncrementables.Text = "Fecha de Pago Incrementables: ";
                rangeFechaIncrementables.Bold = 1;

                rangeFechaIncrementables.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeFechaIncrementables.Text = maskedFechaPagoIncrementables.Text;
                rangeFechaIncrementables.Bold = 0;

                rangeFechaIncrementables.InsertParagraphAfter();

                parrafoEnter.Range.Text = "";
                parrafoEnter.Format.SpaceAfter = 0f;
                parrafoEnter.Range.InsertParagraphAfter();

                objParrafoTitulos.Range.Font.Size = 8;
                objParrafoTitulos.Range.Text = "Decrementables";
                objParrafoTitulos.Range.Bold = 1;
                //Modificar espaciado entre parrafos
                objParrafoTitulos.Format.SpaceAfter = 0f;
                objParrafoTitulos.Format.SpaceBefore = 0f;
                objParrafoTitulos.Range.InsertParagraphAfter();

                //TABLA DECREMENTABLES
                objWord.Range rangetablaDecrementables = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablaDecrementables = objDocument.Tables.Add(rangetablaDecrementables, 1, 10);

                tablaDecrementables.Borders.Enable = 1;

                tamanios.Clear();
                tamanios.AddRange(new float[] { 1.86f, 2.44f, 1.36f, 2.44f, 1.18f, 2.44f, 1.62f, 2.44f, 1.13f, 2.59f });
                tablaDecrementables = asignarTamanioColumna(tablaDecrementables, 10, tamanios);

                celda = tablaDecrementables.Cell(1, 1);
                celda = formatoCelda(celda, "Transporte", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 2);
                celda = formatoCelda(celda, txtboxDreceTransporte.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 3);
                celda = formatoCelda(celda, "Seguro", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 4);
                celda = formatoCelda(celda, txtboxDreceSeguro.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 5);
                celda = formatoCelda(celda, "Carga", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 6);
                celda = formatoCelda(celda, txtboxDreceCarga.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 7);
                celda = formatoCelda(celda, "Embalaje", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 8);
                celda = formatoCelda(celda, txtboxDreceEmbalaje.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 9);
                celda = formatoCelda(celda, "Otros", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaDecrementables.Cell(1, 10);
                celda = formatoCelda(celda, txtboxDreceOtrosGastos.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                objWord.Paragraph fechaDecrementables = objDocument.Paragraphs.Add();
                objWord.Range rangeFechaDecrementables = fechaDecrementables.Range;

                rangeFechaDecrementables.Font.Size = 8;
                rangeFechaDecrementables.Text = "Fecha de Pago Decrementables: ";
                rangeFechaDecrementables.Bold = 1;

                rangeFechaDecrementables.Collapse(objWord.WdCollapseDirection.wdCollapseEnd);

                rangeFechaDecrementables.Text = maskedFechaPagoDecrementables.Text;
                rangeFechaDecrementables.Bold = 0;

                rangeFechaDecrementables.InsertParagraphAfter();

                parrafoEnter.Range.Text = "";
                parrafoEnter.Format.SpaceAfter = 0f;
                parrafoEnter.Range.InsertParagraphAfter();

                //TABLA VALORACION GENERALES DESPACHO ADUANAL
                objWord.Range rangetablasVOGEDA = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablasVOGEDA = objDocument.Tables.Add(rangetablasVOGEDA, 2, 2);

                tablasVOGEDA.Borders.Enable = 1;

                tamanios.Clear();
                tamanios.AddRange(new float[] { 2.41f, 17.17f });
                tablasVOGEDA = asignarTamanioColumna(tablasVOGEDA, 2, tamanios);

                celda = tablasVOGEDA.Cell(1, 1);
                celda = formatoCelda(celda, "Método de Valoración", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablasVOGEDA.Cell(2, 1);
                celda = formatoCelda(celda, "Observaciones Generales de Despacho Aduanal", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablasVOGEDA.Cell(1, 2).Range.Text = cboxMetodoValoracion.Text;
                tablasVOGEDA.Cell(1, 2).Range.Bold = 0;

                tablasVOGEDA.Cell(2, 2).Range.Text = txtboxObservacionesDespachoAduanal.Text;

                parrafoEnter.Range.Text = "";
                parrafoEnter.Format.SpaceAfter = 0f;
                parrafoEnter.Range.InsertParagraphAfter();

                //TABLA FORMA PAGO

                objWord.Range rangetablaFormaPago = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablaFormaPago = objDocument.Tables.Add(rangetablaFormaPago, 2, 3);

                tablaFormaPago.Borders.Enable = 1;

                tamanios.Clear();
                tamanios.AddRange(new float[] { 6.52f, 6.53f, 6.53f });
                tablaFormaPago = asignarTamanioColumna(tablaFormaPago, 3, tamanios);

                celda = tablaFormaPago.Cell(1, 1);
                celda = formatoCelda(celda, "Forma de Pago", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaFormaPago.Cell(2, 1);
                celda = formatoCelda(celda, cboxFormaPago.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaFormaPago.Cell(1, 2);
                celda = formatoCelda(celda, "Condiciones de las Mercancías", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaFormaPago.Cell(2, 2);
                celda = formatoCelda(celda, cboxCondicionesMercancia.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaFormaPago.Cell(1, 3);
                celda = formatoCelda(celda, "Fecha de Pago", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaFormaPago.Cell(2, 3);
                celda = formatoCelda(celda, maskedFechaPagoMercancias.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                parrafoEnter.Range.Text = "";
                parrafoEnter.Format.SpaceAfter = 0f;
                parrafoEnter.Range.InsertParagraphAfter();

                //TABLA RELACIONADO
                objWord.Range rangetablaRelacionado = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablaRelacionado = objDocument.Tables.Add(rangetablaRelacionado, 2, 4);

                tablaRelacionado.Borders.Enable = 1;

                tamanios.Clear();
                tamanios.AddRange(new float[] { 3.49f, 12.5f, 2f , 1.58f});
                tablaRelacionado = asignarTamanioColumna(tablaRelacionado, 4, tamanios);

                celda = tablaRelacionado.Cell(1, 1);
                celda = formatoCelda(celda, "Indicar lo Relacionado al Cumplimiento de NOMs", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaRelacionado.Cell(1, 2).Range.Text = txtNOMs.Text;
                tablaRelacionado.Cell(1, 2).Range.Bold = 0;

                celda = tablaRelacionado.Cell(1, 3);
                celda = formatoCelda(celda, "¿Viene Etiquetado de Origen?", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaRelacionado.Cell(1, 4);
                celda = formatoCelda(celda, cboxEtiquetadoOrigen.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaRelacionado.Cell(2, 1);
                celda = formatoCelda(celda, "¿En Caso de que Requiera Cumplimiento Especificar lo que Corresponda", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaRelacionado.Cell(2, 2).Range.Text = txtboxCumplimiento.Text;
                tablaRelacionado.Cell(2, 2).Range.Bold = 0;

                celda = tablaRelacionado.Cell(2, 3);
                celda = formatoCelda(celda, "¿Utilizarán UVA?", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                celda = tablaRelacionado.Cell(2, 4);
                celda = formatoCelda(celda, cboxUVA.Text, 8, 0, objWord.WdColor.wdColorAutomatic, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                parrafoEnter.Range.Text = "";
                parrafoEnter.Format.SpaceAfter = 0f;
                parrafoEnter.Range.InsertParagraphAfter();

                objParrafoTitulos.Range.Text = "UNA VEZ DESPACHADA LA MERCANCIA ENTREGAR";
                objParrafoTitulos.Range.Font.Size = 9;
                objParrafoTitulos.Range.Bold = 1;
                objParrafoTitulos.Format.SpaceBefore = 0f;
                objParrafoTitulos.Format.SpaceAfter = 0f;
                objParrafoTitulos.Range.InsertParagraphAfter();

                //TABLA UNA VEZ DESPACHADA LA MERCANCIA ENTREGAR
                objWord.Range rangetablaMercanciaDespachada = objDocument.Paragraphs.Add(ref missing).Range;
                objWord.Table tablaMercanciaDespachada = objDocument.Tables.Add(rangetablaMercanciaDespachada, 6, 2);

                tablaMercanciaDespachada.Borders.Enable = 1;

                tamanios.Clear();
                tamanios.AddRange(new float[] { 2.49f, 17.08f });
                tablaMercanciaDespachada = asignarTamanioColumna(tablaMercanciaDespachada, 2, tamanios);

                celda = tablaMercanciaDespachada.Cell(1, 1);
                celda = formatoCelda(celda, "Transporte", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaMercanciaDespachada.Cell(1, 2).Range.Text = txtboxTransporte.Text;
                tablaMercanciaDespachada.Cell(1, 2).Range.Bold = 0;

                celda = tablaMercanciaDespachada.Cell(2, 1);
                celda = formatoCelda(celda, "Tipo Seguro", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaMercanciaDespachada.Cell(2, 2).Range.Text = txtboxTipoSeguro.Text;
                tablaMercanciaDespachada.Cell(2, 2).Range.Bold = 0;

                celda = tablaMercanciaDespachada.Cell(3, 1);
                celda = formatoCelda(celda, "Condiciones Flete", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaMercanciaDespachada.Cell(3, 2).Range.Text = txtboxCondicionesFlete.Text;
                tablaMercanciaDespachada.Cell(3, 2).Range.Bold = 0;

                celda = tablaMercanciaDespachada.Cell(4, 1);
                celda = formatoCelda(celda, "Destino final", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaMercanciaDespachada.Cell(4, 2).Range.Text = txtboxDestinoFinal.Text;
                tablaMercanciaDespachada.Cell(4, 2).Range.Bold = 0;

                celda = tablaMercanciaDespachada.Cell(5, 1);
                celda = formatoCelda(celda, "Custodia", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaMercanciaDespachada.Cell(5, 2).Range.Text = txtboxCustodia.Text;
                tablaMercanciaDespachada.Cell(5, 2).Range.Bold = 0;

                celda = tablaMercanciaDespachada.Cell(6, 1);
                celda = formatoCelda(celda, "Observaciones de Flete a Destino", 8, 1, objWord.WdColor.wdColorGray25, objWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter, objWord.WdParagraphAlignment.wdAlignParagraphCenter);

                tablaMercanciaDespachada.Cell(6, 2).Range.Text = txtboxObservacionesfleteDestino.Text;
                tablaMercanciaDespachada.Cell(6, 2).Range.Bold = 0;

                ruta = rutaCarta.FileName;
                objDocument.SaveAs2(ruta);
                objDocument.Close();
                objAplication.Quit();
                MessageBox.Show("Archivo Generado");
            }


        }

        private objWord.Table asignarTamanioColumna(objWord.Table table, int numColumnas, List<float> tamanios)
        {
            float cmToPoints = 28.35f;
            float[] tamaniosarray = tamanios.ToArray();
            for (int i = 1; i <= numColumnas; i++)
            {
                table.Columns[i].Width = float.Parse(tamaniosarray[i-1].ToString()) * cmToPoints;
            }
            return table;
        }

        private objWord.Cell formatoCelda(objWord.Cell celda,string textocelda, float size , int bold, objWord.WdColor wdColor ,objWord.WdCellVerticalAlignment wdCellVerticalAlignment, objWord.WdParagraphAlignment wdParagraph)
        {
            celda.Range.Text = textocelda;
            celda.Range.Font.Size = size;
            celda.Range.Font.Bold = bold;
            celda.Shading.BackgroundPatternColor = wdColor;
            celda.VerticalAlignment = wdCellVerticalAlignment;
            celda.Range.ParagraphFormat.Alignment = wdParagraph;

            return celda;
        }

        private void dtpFechaPagoIncrementables_CloseUp(object sender, EventArgs e)
        {
            maskedFechaPagoIncrementables.Text = dtpFechaPagoIncrementables.Value.ToString();
        }

        private void dtpFechaPagoDecrementables_CloseUp(object sender, EventArgs e)
        {
            maskedFechaPagoDecrementables.Text = dtpFechaPagoDecrementables.Value.ToString();
        }

        private void dtpFechaPagoMercancias_CloseUp(object sender, EventArgs e)
        {
            maskedFechaPagoMercancias.Text = dtpFechaPagoMercancias.Value.ToString();
        }

        private void txtboxIncreFlete_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxIncreFlete.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxIncreFlete.Text = valor.ToString("G");
            }
        }

        private void txtboxIncreSeguro_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxIncreSeguro.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxIncreSeguro.Text = valor.ToString("G");
            }
        }

        private void txtboxIncreEmbalaje_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxIncreEmbalaje.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxIncreEmbalaje.Text = valor.ToString("G");
            }
        }

        private void txtboxIncreOtrosgastos_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxIncreOtrosgastos.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxIncreOtrosgastos.Text = valor.ToString("G");
            }
        }

        private void txtboxDreceTransporte_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtboxDreceTransporte_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxDreceTransporte.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxDreceTransporte.Text = valor.ToString("G");
            }
        }

        private void txtboxDreceSeguro_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxDreceSeguro.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxDreceSeguro.Text = valor.ToString("G");
            }
        }

        private void txtboxDreceCarga_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxDreceCarga.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxDreceCarga.Text = valor.ToString("G");
            }
        }

        private void txtboxDreceEmbalaje_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxDreceEmbalaje.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxDreceEmbalaje.Text = valor.ToString("G");
            }
        }

        private void txtboxDreceOtrosGastos_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxDreceOtrosGastos.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxDreceOtrosGastos.Text = valor.ToString("G");
            }
        }

        private void txtboxValorFactura1_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxDreceOtrosGastos.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxDreceOtrosGastos.Text = valor.ToString("G");
            }
        }

        private void txtboxValorFactura2_Enter(object sender, EventArgs e)
        {
            // Al entrar, quitamos los separadores de miles para que sea fácil editar
            if (decimal.TryParse(txtboxValorFactura2.Text, out decimal valor))
            {
                // "G" es formato general, mantiene los decimales necesarios sin ceros extra
                txtboxValorFactura2.Text = valor.ToString("G");
            }
        }

        private void gboxImportador_Enter(object sender, EventArgs e)
        {

        }

        private void panelCartanInstrucciones_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
