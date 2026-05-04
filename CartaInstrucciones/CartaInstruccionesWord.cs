using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using CartaInstrucciones.Modelo;
using objWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;
using SixLabors.ImageSharp.Metadata.Profiles.Exif;
using Google.Protobuf.Compiler;

namespace CartaInstrucciones
{
    internal class CartaInstruccionesWord
    {
        private static CartaInstruccionesWord _intancia = null;

        public CartaInstruccionesWord() { }
        public static CartaInstruccionesWord Instancia
        {
            get
            {
                if (_intancia == null)
                {
                    _intancia = new CartaInstruccionesWord();
                }

                return _intancia;
            }
        }

        public bool generarCartaInstrucciones()
        {
            bool generar = false;
            
            SaveFileDialog rutaCarta = new SaveFileDialog();

            rutaCarta.Filter = "Archivos de texto|*.txt|Todos los archivos|*.*";
            rutaCarta.Title = "Guardar Carta de Instrucciones";

            rutaCarta.ShowDialog();
            string ruta;
            if (rutaCarta.FileName != "")
            {                  
                objWord.Application objAplication = new objWord.Application();

                objWord.Document objDocument = objAplication.Documents.Add();

                objWord.PageSetup pageSetup = objDocument.PageSetup;

                float cmToPoints = 28.35f;
                pageSetup.TopMargin = 1.59f * cmToPoints;
                pageSetup.BottomMargin = 2.22f * cmToPoints;
                pageSetup.LeftMargin = 2.54f * cmToPoints;
                pageSetup.RightMargin = 2.54f * cmToPoints;


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

                //TAMAÑO DE LETRA
                objParrafo1.Range.Font.Size = 12;
                //AGREGAR TEXTO A UN PARRAFO
                objParrafo1.Range.Text = "OTRO PARRAFO SIENDO EL MISMO PARRAFO";
                //TIPO DE LETRA
                objParrafo1.Range.Font.Name = "Arial";
                //CENTRAR UN PARRAFO
                objParrafo1.Alignment = objWord.WdParagraphAlignment.wdAlignParagraphCenter;

                objParrafo1.Range.Font.Size = 8;

                ruta = rutaCarta.FileName;
            

                objParrafo1.Range.InsertParagraphAfter();


                objDocument.SaveAs2(ruta);
                generar = true;
                objDocument.Close();
                objAplication.Quit();
                
            }
            return generar;
        }
    }
}
