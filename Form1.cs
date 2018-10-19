using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using Microsoft.Office.Interop.Word;





namespace GeneradorConstancias
{
    
    

    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application _excelApp;
        Workbook book;
        Workbook bookWord;
        string nombreArchivo;
        string nombreArchivoWord;
        Microsoft.Office.Interop.Word.Application _wordApp;

        int progreso;
        InfoComun infocomun = new InfoComun();
        List<Alumno> alumnos=  new List<Alumno>();
        public Form1()
        {
            InitializeComponent();
        }

        private void btnAbrirExcel_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result==DialogResult.OK)
            {
                this.nombreArchivo = openFileDialog1.FileName;
                txtRuta.Text = this.nombreArchivo;
                habilitarBoton();
            }
         
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            _excelApp = new Microsoft.Office.Interop.Excel.Application();
            book = _excelApp.Workbooks.Add(this.nombreArchivo);
            Worksheet sheet = (Worksheet)book.Sheets[1];
            Worksheet sheet2 = (Worksheet)book.Sheets[2];

            
            infocomun.Carrera = sheet2.Cells[2, 1].value;
            infocomun.Grado_asesor = sheet2.Cells[2, 2].value;
            //MessageBox.Show(infocomun.Grado_asesor);
            infocomun.Nombre_asesor= sheet2.Cells[2, 3].value;
            infocomun.Ciclo = sheet2.Cells[2, 4].value;
            infocomun.Fecha = sheet2.Cells[2, 5].value;
            sheet.Rows.ClearFormats();
            for (int x=2;x<=sheet.UsedRange.Rows.Count;x++)
            {
                
                
                   // MessageBox.Show(sheet.Cells[x, 1].value);
                    Alumno alumno = new Alumno();
                    alumno.Nombre = sheet.Cells[x, 1].value;
                    alumno.Proyecto = sheet.Cells[x, 2].value;
                    //alumno.Lugar = sheet.Cells[x, 3].value;
                    //alumno.Fecha = sheet.Cells[x, 4].value;
                  //  MessageBox.Show(alumno.Proyecto);
                    alumnos.Add(alumno);
                

            }
            actualizarProgreso(30, "leyendo excel..");

            book.Close();

            generarConstancias();  
         
        }

           public void actualizarProgreso(int valor, string comentario)
        {
            progreso += valor;
            System.Threading.Thread.Sleep(1000);
            progressBar1.Value = progreso;
            lblProgreso.Text = comentario;
            if (progressBar1.Value == 100)
            {
                MessageBox.Show("CONSTANCIAS GENERADAS");
                progressBar1.Value = 0;
                lblProgreso.Text = "Progreso";
                progreso = 0;
                alumnos.Clear();
            }
        }
        private void habilitarBoton()
        {
            if(txtRuta.Text!="" && txtRutaWord.Text != "")
            {
                btnGenerar.Enabled = true;
            }
        }
        private void generarConstancias()
        {
            foreach( var alumno in alumnos)
            {
                _wordApp = new Microsoft.Office.Interop.Word.Application();
                Document wordDoc = _wordApp.Documents.Add(this.nombreArchivoWord);
                Bookmark nombre_alumno = wordDoc.Bookmarks["nombre_alumno"];
                Bookmark nombre_proyecto = wordDoc.Bookmarks["nombre_proyecto"];
                Bookmark nombre_asesor = wordDoc.Bookmarks["nombre_asesor"];
                Bookmark grado_asesor = wordDoc.Bookmarks["grado_asesor"];
                Bookmark fecha = wordDoc.Bookmarks["fecha"];
                Bookmark ciclo = wordDoc.Bookmarks["ciclo"];
                Bookmark carrera = wordDoc.Bookmarks["carrera"];
                Microsoft.Office.Interop.Word.Range nombre_alumno_r = nombre_alumno.Range;
                Microsoft.Office.Interop.Word.Range nombre_proyecto_r = nombre_proyecto.Range;
                Microsoft.Office.Interop.Word.Range nombre_asesor_r = nombre_asesor.Range;
                Microsoft.Office.Interop.Word.Range grado_asesor_r = grado_asesor.Range;
                Microsoft.Office.Interop.Word.Range fecha_r = fecha.Range;
                Microsoft.Office.Interop.Word.Range ciclo_r = ciclo.Range;
                Microsoft.Office.Interop.Word.Range carrera_r = carrera.Range;

                nombre_alumno_r.Text = alumno.Nombre;
                nombre_proyecto_r.Text = alumno.Proyecto;
                grado_asesor_r.Text = infocomun.Grado_asesor;
                nombre_asesor_r.Text = infocomun.Nombre_asesor;
                fecha_r.Text = infocomun.Fecha;
                ciclo_r.Text = infocomun.Ciclo;
                carrera_r.Text = infocomun.Carrera;
                string nomArchivoNuevo;
                nomArchivoNuevo = alumno.Nombre.ToString().Trim() + "_" + infocomun.Ciclo.ToString().Trim() + ".docx";
               // MessageBox.Show(nomArchivoNuevo);
               
                wordDoc.SaveAs2(@"C:\Constancias\" + nomArchivoNuevo);
                wordDoc.Close();


            }
            actualizarProgreso(70, "generando constancias...");
        }

        private void btn_abrirWord_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.nombreArchivoWord = openFileDialog2.FileName;
                txtRutaWord.Text = this.nombreArchivoWord;
                habilitarBoton();
            }
          
        }
    }
}
public class Alumno
{

    string nombre;
    string proyecto;
   // string lugar;
    //string fecha;

    public string Nombre { get => nombre; set => nombre = value; }
    public string Proyecto { get => proyecto; set => proyecto = value; }
   // public string Lugar { get => lugar; set => lugar = value; }
    //public string Fecha { get => fecha; set => fecha = value; }
}

public class InfoComun
{

    string carrera;
    string grado_asesor;
    string nombre_asesor;
    string fecha;
    string ciclo;

    public string Carrera { get => carrera; set => carrera = value; }
    public string Grado_asesor { get => grado_asesor; set => grado_asesor = value; }
    public string Nombre_asesor { get => nombre_asesor; set => nombre_asesor = value; }
    public string Fecha { get => fecha; set => fecha = value; }
    public string Ciclo { get => ciclo; set => ciclo = value; }
}