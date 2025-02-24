using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using GeneradorNotificacionesPreJurídicos.Models;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Microsoft.Office.Interop.Word;
using System.Globalization;


namespace GeneradorNotificacionesPreJurídicos
{
    public partial class Form1 : Form
    {
        Repository repository = new Repository();
        Utility utility = new Utility();
        private OpenFileDialog openFileDialog1;
        public Form1()
        {
            InitializeComponent();
            openFileDialog1 = new OpenFileDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ResetRadioButtons();
            openFileDialog1.Filter = "Archivos de Excel (*.xlsx)|*.xlsx";
            openFileDialog1.Title = "Seleccionar archivo de Excel";
            dtpFecha.Format = DateTimePickerFormat.Custom;
            dtpFecha.CustomFormat = "dd MMMM yyyy";

        }

        private void ResetRadioButtons()
        {

            rbtnGrupoClaves.Checked = true;
        }

        private void rbtnClave_CheckedChanged(object sender, EventArgs e)
        {
     
            grpBoxArchivo.Visible = true;
            txtArchivo.Clear();
        }

        private void rbtnGrupoClaves_CheckedChanged(object sender, EventArgs e)
        {
            grpBoxArchivo.Visible = true;
          
        }

        private void btnSeleccionarArchivo_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtArchivo.Text = openFileDialog1.FileName;
                btnProcesarArchivo.Focus();
            }
        }

        private void txtArchivo_TextChanged(object sender, EventArgs e)
        {
            if (txtArchivo.Text.Length > 0)
            {
                btnProcesarArchivo.Enabled = true;
            }
        }

        private void btnProcesarArchivo_Click(object sender, EventArgs e)
        {
            if (txtArchivo.Text.Length > 0)
            {
                Cursor.Current = Cursors.WaitCursor;

                string filePath = txtArchivo.Text;
                int columnNumber = 1;

               

                if (rbtnGrupoClaves.Checked)
                {
                    using (var workbook = new XLWorkbook(filePath))
                    {
                        var worksheet = workbook.Worksheet(1); // Obtener la primera hoja del libro
                        List<string> claves = new List<string>();
                        List<string> Gestores = new List<string>();
                    

                        foreach (var cell in worksheet.Column(columnNumber).CellsUsed())
                        {
                            if(int.TryParse(cell.Value.ToString().Trim(), out int result))
                            {
                                claves.Add(cell.Value.ToString().Trim());
                            }
                        }

                        // Leer la segunda columna para 'gestores'
                        var gestorCells = worksheet.Column(2).CellsUsed().Skip(1);
                        foreach (var cell in gestorCells)
                        {
                            string cellValue = cell.Value.ToString().Trim();
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                Gestores.Add(cellValue);
                            }
                        }

                        // Crear el diccionario para mapear claves con gestores
                        Dictionary<string, string> claveGestorMap = new Dictionary<string, string>();
                        for (int i = 0; i < claves.Count; i++)
                        {
                            if (i < Gestores.Count)
                            {
                                claveGestorMap[claves[i]] = Gestores[i];
                            }
                        }



                        if (claves.Any())
                        {
                            string response = "";

                            var informacionCliente = repository.GetInformacionClientes(claves);
                            var nisRads = informacionCliente.Select(x => x.NisRad.ToString()).ToList();
                            var areas = informacionCliente.Select(x => x.Area.ToString()).Distinct().ToList();

                            var SaldoFinanciado = repository.GetSaldoFinanciados(nisRads);
                            //var DeudaCliente = repository.GetDeudaClientes(informacionCliente.FirstOrDefault().NisRad);

                            
                            DateTime fecha = dtpFecha.Value;

                            
                            int dia = fecha.Day;
                            string mes = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(fecha.Month);
                            mes = char.ToUpper(mes[0]) + mes.Substring(1);
                            int anio = fecha.Year;

                            //// Construir la cadena formateada
                            //string fechaFormateada = $"{dia} de {mes} del {anio}";


                            var InfoLider = repository.GetInfoLiders(areas);
                            var InfoGestor = repository.GetInfoGestor(Gestores);
                           

                            if (informacionCliente.Count > 0)
                            {
                                response = utility.GenerateNotificaciones(informacionCliente,  SaldoFinanciado, InfoLider,InfoGestor, dia, mes,anio, claveGestorMap);
                            }
                        }
                        else
                        {
                            //Console.WriteLine($"El valor '{cell.Value.ToString().Trim()}' no es un número entero válido.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No es un valor entero:");
                }

                txtArchivo.Clear();
                Cursor.Current = Cursors.Default;

                MessageBox.Show("Se generó las notificaciones proporcionadas", "Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}
