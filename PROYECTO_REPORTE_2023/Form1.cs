using BE;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace PROYECTO_REPORTE_2023
{

    public partial class Form1 : Form
    {
        List<string> Start = new List<string>();
        List<string> End = new List<string>();
        BE_Persona bePersona;
        BE_Horario beHorario;

        List<BE_Persona> ListPerson = new List<BE_Persona>();
        public Form1()
        {
            InitializeComponent();
            CenterToScreen();
            this.txtRutaArchivo.Enabled = false;
        }

        protected string NombreArchivo(string cad)
        {
            DateTime dt = DateTime.Now;
            string info = dt.ToShortTimeString();
            string infoReal = info.Replace(" ", "");
            string inforReal2 = infoReal.Replace(":", "-");
            string ruta = Path.GetDirectoryName(cad);
            string nombreArchivo = @"" + ruta + "\\Reporte" + inforReal2 + ".xls";
            return nombreArchivo;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnSeleccionar_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.Filter = "Archivos Excel|*.xlsx;*.xls";
            if (oFD.ShowDialog() == DialogResult.OK)
            {
                string filePath = oFD.FileName;
                txtRutaArchivo.Text = filePath;
            }
        }

        protected void CerrarHoja(Excel.Workbook wk, Excel.Worksheet ws, Excel.Application app)
        {
            wk.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wk);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
        private void btnGenerar_Click(object sender, EventArgs e)
        {
            System.Drawing.Font boldFont = new System.Drawing.Font(btnGenerar.Font, FontStyle.Bold);
            btnGenerar.Font = boldFont;
            btnGenerar.Enabled = false;
            btnGenerar.Text = "Generando.";

            //RECONOCIMIENTO DE CASILLAS
            string ruta = txtRutaArchivo.Text;
            var primerArchivo = new Excel.Application();
            Excel.Workbook workbook = primerArchivo.Workbooks.Open(@"" + ruta);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            Excel.Range usedRange = worksheet.UsedRange.Columns["A"];
            Excel.Range searchRange = usedRange.Cells;
            foreach (Excel.Range cell in searchRange)
            {
                if (cell.Value != null && cell.Value.ToString().Contains("Lun"))
                {
                    Excel.Range inicioEncabezado = worksheet.Range[cell.Address.ToString()];
                    Start.Add(cell.Address.ToString());
                }
                int numericValue;
                if (cell.Value != null && int.TryParse(cell.Value.ToString(), out numericValue))
                {
                    Excel.Range finPieCuadro = worksheet.Range[cell.Address.ToString()];
                    End.Add(cell.Address.ToString());
                }
            }
            //VALIDACION DE DATOS
            if(Start.Count==0 && End.Count== 0)
            {
                MessageBox.Show("Seleccione un reporte valido para el programa");
                return;
            }

            //CREACION DE HOJA DE EXCEL
            Excel.Application segundoArchivo = new Excel.Application();
            Excel.Workbook workbook2 = segundoArchivo.Workbooks.Add();
            Excel.Worksheet worksheet2 = workbook2.ActiveSheet;
            worksheet2.Name = "Hoja1";
            Excel.Range columnNombre = worksheet2.Columns[2];
            Excel.Range columnFecha1 = worksheet2.Columns[3];
            Excel.Range columnFecha2 = worksheet2.Columns[4];
            columnNombre.ColumnWidth = 35;
            columnFecha1.ColumnWidth = 16;
            columnFecha2.ColumnWidth = 16;
            worksheet2.Cells[1, 2] = "NOMBRE";
            worksheet2.Cells[1, 3] = "ENTRADA";
            worksheet2.Cells[1, 4] = "SALIDA";
            Excel.Range usedRange2 = worksheet2.UsedRange;
            btnGenerar.Text = "Generando..";
            //GET DE VALORES
            for (int i = 0; i<Start.Count; i++)
            {
                Excel.Range valorCelda = worksheet.Range[End[i]];
                bePersona = new BE_Persona();
                bePersona.nombre = valorCelda.Offset[0, 1].Value; //valor de la celda

                Excel.Range CoordenadaInicio = worksheet.Range[Start[i]];
                int fila = CoordenadaInicio.Row;
                int columna = CoordenadaInicio.Column;

                Excel.Range CoordenadaFinal = worksheet.Range[End[i]];
                int filaFinal = CoordenadaFinal.Row;
                List<BE_Horario> listaHorarios = new List<BE_Horario>();
                for (int j = fila; j<filaFinal; j++)
                {
                    beHorario = new BE_Horario();
                    Excel.Range CeldaFecha = worksheet.Cells[j,columna  + 1];
                    Excel.Range primeraMarca = worksheet.Cells[j,columna + 4];
                    if (primeraMarca.Value == null)
                    {
                        beHorario.HoraEntrada = CeldaFecha.Value.ToString() + " - ";
                    }
                    else
                    {
                        beHorario.HoraEntrada = CeldaFecha.Value.ToString() + " " + getHora(primeraMarca);                               
                    }
                    
                    Excel.Range segundaMarca = worksheet.Cells[j, columna + 5];

                    if (segundaMarca.Value == null)
                    {
                        beHorario.HoraSalida = CeldaFecha.Value.ToString() + " - ";
                    }else
                        {
                        beHorario.HoraSalida = CeldaFecha.Value.ToString() + " " + getHora(segundaMarca);
                        }
                    listaHorarios.Add(beHorario);
                }
                bePersona.ListHorario = listaHorarios;
                ListPerson.Add(bePersona);
            }
            

            foreach (BE_Persona perso in ListPerson)
            {
                foreach (BE_Horario horario in perso.ListHorario)
                {
                    int lastRow = usedRange2.Rows.Count + 1;
                    Excel.Range row = worksheet2.Rows[lastRow];
                    row.Insert();
                    worksheet2.Cells[lastRow, 2].Value = perso.nombre;
                    worksheet2.Cells[lastRow, 3].Value = horario.HoraEntrada;
                    worksheet2.Cells[lastRow, 4].Value = horario.HoraSalida;
                }
            }
            //edicion de celdas
            btnGenerar.Text = "Generando...";
            //color de cabeza
            Excel.Range rangeColorear = worksheet2.Range["B1:D1"];
            rangeColorear.Font.Bold = true;
            rangeColorear.Interior.Color = System.Drawing.Color.FromArgb(235, 231, 230).ToArgb();
            //bordes
            Excel.Range usedRange3 = worksheet2.UsedRange;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            usedRange3.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            
            MessageBox.Show("Se genero exitosamente el reporte!");
            string cadena = txtRutaArchivo.Text;
            workbook2.SaveAs(NombreArchivo(cadena));
            workbook2.Close();
            segundoArchivo.Quit();
            //mesaje de exito
            
            CerrarHoja(workbook, worksheet, primerArchivo);
            System.Drawing.Font normalFont = new System.Drawing.Font(btnGenerar.Font, FontStyle.Regular);
            btnGenerar.Enabled = true;
            btnGenerar.Font = normalFont;
            btnGenerar.Text = "Generar";
            txtRutaArchivo.Text = "";

        }
        protected string getHora(Excel.Range horaObtenida) {

            string segundaMarcaString = Convert.ToString(horaObtenida.Value);
            if (double.TryParse(segundaMarcaString, out double hora))
            {
                TimeSpan tiempo = TimeSpan.FromHours(hora * 24);
                int horas = tiempo.Hours;
                int minutos = tiempo.Minutes;

                string horaFormateada = $"{horas:D2}:{minutos:D2}";
                return horaFormateada;
            }
            else
            {
                return horaObtenida.Value;
            }
        }


    }
}
