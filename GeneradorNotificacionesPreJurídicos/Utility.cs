using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.ExtendedProperties;
using GeneradorNotificacionesPreJurídicos.Models;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using _Application = Microsoft.Office.Interop.Word._Application;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using System.Security.Policy;
using DocumentFormat.OpenXml.InkML;
using System.Numerics;
//using DocumentFormat.OpenXml.EMMA;



namespace GeneradorNotificacionesPreJurídicos
{
    public class Utility
    {
        Repository repository = new Repository();

        string directorioDestino = $@"C:\Notificaciones Pre-Jurídico\{DateTime.Now.ToString("dd-MM-yyyy")}";
        public string GenerateNotificaciones(List<InformacionCliente> informacionCliente, List<SaldoFinanciado> SaldoFinanciado, List<InfoLider> InfoLider, List<InfoGestor> InfoGestor, int dia, string mes, int anio, Dictionary<string, string> claveGestorMap)
        {

            var leftJoinInformacion = from ic in informacionCliente
                                      join sf in SaldoFinanciado on ic.NisRad equals sf.Nisrad into sfJoin
                                      from sf in sfJoin.DefaultIfEmpty()
                                      join il in InfoLider on ic.Area.ToUpper().Trim() equals il.SectorL.ToUpper().Trim() into ilJoin
                                      from il in ilJoin.DefaultIfEmpty()
                                      join ig in InfoGestor on il.SectorL.ToUpper().Trim() equals ig.SectorG.ToUpper().Trim() into igJoin
                                      from ig in igJoin.DefaultIfEmpty()
                                      join cg in claveGestorMap on ic.Clave.ToUpper().Trim() equals cg.Key.ToUpper().Trim() into cgJoin
                                      from cg in cgJoin.DefaultIfEmpty()  // <== corregido
                                      group new { ic, sf, il, ig, cg } by new
                                      {
                                          ic.Clave,
                                          ic.NisRad,
                                          ic.NombreCliente,
                                          ic.DireccionCliente,
                                          ic.Area,
                                          ic.FechaUltimoPago,
                                          ic.DeudaTotal,
                                          ic.FechaActual,
                                          ic.DiaActual,
                                          ic.MesActual,
                                          ic.AnioActual,
                                          ic.MesActualText,
                                          ic.DiaActualText
                                      } into grouped
                                      select new
                                      {
                                          Clave = grouped.Key.Clave,
                                          NisRad = grouped.Key.NisRad,
                                          NombreCliente = grouped.Key.NombreCliente,
                                          DireccionCliente = grouped.Key.DireccionCliente,
                                          Area = grouped.Key.Area,
                                          FechaUltimoPago = grouped.Key.FechaUltimoPago,
                                          DeudaTotal = grouped.Key.DeudaTotal,
                                          FechaActual = grouped.Key.FechaActual,
                                          DiaActual = grouped.Key.DiaActual,
                                          MesActual = grouped.Key.MesActual,
                                          AnioActual = grouped.Key.AnioActual,
                                          MesActualText = grouped.Key.MesActualText,
                                          DiaActualText = grouped.Key.DiaActualText,
                                          SaldoF = grouped.Select(g => g.sf != null ? g.sf.SaldoF : 0).FirstOrDefault(),
                                          SectorL = grouped.Select(g => g.il != null ? g.il.SectorL : null).FirstOrDefault(),
                                          NombreL = grouped.Select(g => g.il != null ? g.il.NombreL : null).FirstOrDefault(),
                                          AnalistC = grouped.Select(g => g.il != null ? g.il.AnalistaC : null).FirstOrDefault(),
                                          Telefono = grouped.Select(g => g.il != null ? g.il.Telefono : null).FirstOrDefault(),
                                          Email = grouped.Select(g => g.il != null ? g.il.Email : null).FirstOrDefault(),
                                          FirmaElectronica = grouped.Select(g => g.il != null ? g.il.FirmaElectronica : null).FirstOrDefault(),
                                          SectorG = grouped.Select(g => g.ig != null ? g.ig.SectorG : null).FirstOrDefault(),
                                          Usuario = grouped.Select(g => g.ig != null ? g.cg.Value : null).FirstOrDefault(),
                                          //nombreg = grouped.Where(g => g.ig != null && !string.IsNullOrEmpty(g.cg.Value) && g.ig.Usuario == g.cg.Value).Select(g => g.ig.NombreG).FirstOrDefault(),
                                          nombreg = grouped.Where(g => g.ig != null && !string.IsNullOrEmpty(g.cg.Value) && g.ig.Usuario.Trim().ToUpper() == g.cg.Value.Trim().ToUpper()).Select(g => g.ig.NombreG).FirstOrDefault(),
                                          Telefonog = grouped.Where(g => g.ig != null && !string.IsNullOrEmpty(g.cg.Value) && g.ig.Usuario.Trim().ToUpper() == g.cg.Value.Trim().ToUpper()).Select(g => g.ig.Telefono).FirstOrDefault(),
                                          //Telefonog = grouped.Select(g => g.ig != null ? g.ig.Telefono : null).FirstOrDefault(),
                                          Cargo = grouped.Select(g => g.ig != null ? g.ig.Cargo : null).FirstOrDefault(),
                                      };




            List<string> archivos = new();

            foreach (var item in leftJoinInformacion)
            {
                try
                {
                    KillProcess("WINWORD.EXE");
                    KillProcess("Acrord32.exe");

                    if (!Directory.Exists(directorioDestino))
                    {
                        Directory.CreateDirectory(directorioDestino);
                    }

                    System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("es-HN");

                    // Creando documento de Word
                    Word.Application wordApp = new Word.Application();
                    Word.Document document = wordApp.Documents.Open($@"{Environment.CurrentDirectory}/plantilla notificaciones de corte.docx");
                    Word.Range range = document.Content;
                    object missing = System.Reflection.Missing.Value;


                    //ENCABEZADO
                    Word.Paragraph parrafoEncabezado = document.Content.Paragraphs.Add(ref missing);
                    parrafoEncabezado.Range.Text = $@"";
                    parrafoEncabezado.Range.Font.Size = 12;
                    parrafoEncabezado.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoEncabezado.Range.Font.Bold = 0;

                    Word.Range parrafoValorEncabezado = parrafoEncabezado.Range;
                    parrafoValorEncabezado.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoValorEncabezado.Text = $@"{item.Area}, {dia} de {mes} del {anio}{Environment.NewLine}";
                    parrafoValorEncabezado.Font.Size = 12;
                    parrafoValorEncabezado.Font.Name = "Calibri (Cuerpo)";
                    parrafoValorEncabezado.Font.Bold = 0;

                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    parrafoEncabezado.Range.InsertParagraphAfter();

                    // NOMBRE
                    Word.Paragraph parrafoNombre = document.Content.Paragraphs.Add(ref missing);
                    parrafoNombre.Range.Text = $@"Nombre de Cliente: ";
                    parrafoNombre.Range.Font.Size = 12;
                    parrafoNombre.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoNombre.Range.Font.Bold = 0;

                    Word.Range parrafoValorNombre = parrafoNombre.Range;
                    parrafoValorNombre.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoValorNombre.Text = $@"{item.NombreCliente}";
                    parrafoValorNombre.Font.Size = 12;
                    parrafoValorNombre.Font.Name = "Calibri (Cuerpo)";
                    parrafoValorNombre.Font.Bold = 1;

                    parrafoNombre.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    parrafoNombre.Range.InsertParagraphAfter();

                    // CLAVE
                    Word.Paragraph parrafoClave = document.Content.Paragraphs.Add(ref missing);
                    parrafoClave.Range.Text = $@"Clave: ";
                    parrafoClave.Range.Font.Size = 12;
                    parrafoClave.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoClave.Range.Font.Bold = 0;

                    Word.Range parrafoValorClave = parrafoClave.Range;
                    parrafoValorClave.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoValorClave.Text = $@"{item.Clave}";
                    parrafoValorClave.Font.Size = 12;
                    parrafoValorClave.Font.Name = "Calibri (Cuerpo)";
                    parrafoValorClave.Font.Bold = 1;

                    parrafoClave.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    parrafoClave.Range.InsertParagraphAfter();


                    // DIRECCION
                    Word.Paragraph parrafoDireccion = document.Content.Paragraphs.Add(ref missing);
                    parrafoDireccion.Range.Text = $@"Dirección: ";
                    parrafoDireccion.Range.Font.Size = 12;
                    parrafoDireccion.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoDireccion.Range.Font.Bold = 0;

                    Word.Range parrafoValorDireccion = parrafoNombre.Range;
                    parrafoValorDireccion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoValorDireccion.Text = $@"{item.DireccionCliente}{Environment.NewLine}";
                    parrafoValorDireccion.Font.Size = 12;
                    parrafoValorDireccion.Font.Name = "Calibri (Cuerpo)";
                    parrafoValorDireccion.Font.Bold = 1;

                    parrafoDireccion.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    parrafoDireccion.Range.InsertParagraphAfter();

                    //ASUNTO
                    Word.Paragraph parrafoAsunto = document.Content.Paragraphs.Add(ref missing);
                    parrafoAsunto.Range.Text = $@"ASUNTO: NOTIFICACIÓN DE DEUDA Y REQUERIMIENTO DE PAGO";
                    parrafoAsunto.Range.Font.Size = 12;
                    parrafoAsunto.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoAsunto.Range.Font.Bold = 1;

                    parrafoAsunto.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    parrafoAsunto.Range.InsertParagraphAfter();

                    //Descripcion
                    Word.Paragraph parrafoDescripcion = document.Content.Paragraphs.Add(ref missing);
                    parrafoDescripcion.Range.Text = $@"La Empresa Nacional de Energía Eléctrica (ENEE) le informa que, según nuestros registros, su cuenta presenta un saldo en mora bajo la clave ";
                    parrafoDescripcion.Range.Font.Size = 12;
                    parrafoDescripcion.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoDescripcion.Range.Font.Bold = 0;
                    parrafoDescripcion.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;

                    //CLAVE
                    Word.Range parrafoCla = parrafoDescripcion.Range;
                    parrafoCla.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoCla.Text = $@"{item.Clave}";
                    parrafoCla.Font.Size = 12;
                    parrafoCla.Font.Name = "Calibri (Cuerpo)";
                    parrafoCla.Font.Bold = 1;

                    //Sigue descripción
                    Word.Range parrafoDes1 = parrafoDescripcion.Range;
                    parrafoDes1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes1.Text = $@". A la fecha, el saldo adeudado asciende a ";
                    parrafoDes1.Font.Size = 12;
                    parrafoDes1.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes1.Font.Bold = 0;

                    parrafoDireccion.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;

                    //DEUDA
                    var deudaTotal = item.DeudaTotal;
                    string deudaFormateada = deudaTotal.ToString("N2");

                    Word.Range parrafoDeuda = parrafoDescripcion.Range;
                    parrafoDeuda.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDeuda.Text = $@"L. {deudaFormateada}";
                    parrafoDeuda.Font.Size = 12;
                    parrafoDeuda.Font.Name = "Calibri (Cuerpo)";
                    parrafoDeuda.Font.Bold = 1;

                    //if (item.SaldoF > 0)
                    //{
                    //Sigue descripción
                    Word.Range parrafoDes2 = parrafoDescripcion.Range;
                    parrafoDes2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes2.Text = $@", además de un monto adicional de  ";
                    parrafoDes2.Font.Size = 12;
                    parrafoDes2.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes2.Font.Bold = 0;

                    parrafoDireccion.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;

                    //SALDOFINACIADO
                    decimal Saldo = item.SaldoF;
                    string saldoFormateada = Saldo.ToString("N2");

                    Word.Range parrafoSaldo = parrafoDescripcion.Range;
                    parrafoSaldo.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoSaldo.Text = $@"L. {saldoFormateada} ";
                    parrafoSaldo.Font.Size = 12;
                    parrafoSaldo.Font.Name = "Calibri (Cuerpo)";
                    parrafoSaldo.Font.Bold = 1;

                    //Sigue descripción
                    Word.Range parrafoDes21 = parrafoDescripcion.Range;
                    parrafoDes21.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes21.Text = $@"correspondiente a un acuerdo de pago previamente establecido.  ";
                    parrafoDes21.Font.Size = 12;
                    parrafoDes21.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes21.Font.Bold = 0;

                    parrafoDireccion.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;

                    //}

                    //Sigue descripción
                    Word.Range parrafoDes3 = parrafoDescripcion.Range;
                    parrafoDes3.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes3.Text = $@"{Environment.NewLine}Es fundamental que atienda este asunto de inmediato. Le solicitamos realizar el pago total de su deuda, correspondiente al suministro de energía eléctrica proporcionado por la ENEE. Entendemos que en ocasiones pueden surgir imprevistos; por ello, si no es posible cancelar la totalidad del monto, el gestor de cobro que le entrega esta notificación está autorizado para ofrecerle opciones de pago, como un arreglo de pago o una autorización para un pago parcial, en línea con nuestras políticas. ";
                    parrafoDes3.Font.Size = 12;
                    parrafoDes3.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes3.Font.Bold = 0;

                    Word.Range parrafoDes5 = parrafoDescripcion.Range;
                    parrafoDes5.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes5.Text = $@"{Environment.NewLine}Le instamos a que, dentro de las próximas ";
                    parrafoDes5.Font.Size = 12;
                    parrafoDes5.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes5.Font.Bold = 0;

                    Word.Range parrafoDes6 = parrafoDescripcion.Range;
                    parrafoDes6.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes6.Text = "48 horas ";
                    parrafoDes6.Font.Size = 12;
                    parrafoDes6.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes6.Font.Bold = 1;

                    Word.Range parrafoDes7 = parrafoDescripcion.Range;
                    parrafoDes7.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes7.Text = $@"a partir de la recepción de este documento, realice su pago. De no realizarlo o establecer un acuerdo para regularizar su situación, la ENEE se verá en la obligación de remitir su cuenta a gestión de ";
                    parrafoDes7.Font.Size = 12;
                    parrafoDes7.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes7.Font.Bold = 0;

                    Word.Range parrafoDes61 = parrafoDescripcion.Range;
                    parrafoDes61.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes61.Text = "cobro jurídico";
                    parrafoDes61.Font.Size = 12;
                    parrafoDes61.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes61.Font.Bold = 1;

                    Word.Range parrafoDes71 = parrafoDescripcion.Range;
                    parrafoDes71.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes71.Text = $@", lo que podría implicar procesos a través de los Tribunales de la República. ";
                    parrafoDes71.Font.Size = 12;
                    parrafoDes71.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes71.Font.Bold = 0;


                    Word.Range parrafoDes8 = parrafoDescripcion.Range;
                    parrafoDes8.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes8.Text = $@"{Environment.NewLine}Para cualquier consulta adicional o para concretar un acuerdo de pago, puede comunicarse directamente con su gestor de cobro al número ";
                    parrafoDes8.Font.Size = 12;
                    parrafoDes8.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes8.Font.Bold = 0;

                    Word.Range parrafoDes10 = parrafoDescripcion.Range;
                    parrafoDes10.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes10.Text = $@"{item.Telefonog} ";
                    parrafoDes10.Font.Size = 12;
                    parrafoDes10.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes10.Font.Bold = 1;

                    Word.Range parrafoDes11 = parrafoDescripcion.Range;
                    parrafoDes11.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes11.Text = $@"o con nuestro equipo de backoffice al número ";
                    parrafoDes11.Font.Size = 12;
                    parrafoDes11.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes11.Font.Bold = 0;

                    Word.Range parrafoDes12 = parrafoDescripcion.Range;
                    parrafoDes12.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes12.Text = $@"{item.Telefono}";
                    parrafoDes12.Font.Size = 12;
                    parrafoDes12.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes12.Font.Bold = 1;

                    Word.Range parrafoDes111 = parrafoDescripcion.Range;
                    parrafoDes111.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes111.Text = $@" Estamos aquí para ayudarle de forma rápida y efectiva. ";
                    parrafoDes111.Font.Size = 12;
                    parrafoDes111.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes111.Font.Bold = 0;

                    Word.Range parrafoDes101 = parrafoDescripcion.Range;
                    parrafoDes101.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes101.Text = $@"{Environment.NewLine} Atentamente, ";
                    parrafoDes101.Font.Size = 12;
                    parrafoDes101.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes101.Font.Bold = 0;

                    parrafoDireccion.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;


                    Word.Range parrafoDes112 = parrafoDescripcion.Range;
                    parrafoDes112.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoDes112.Text = $@"{Environment.NewLine}";
                    parrafoDes112.Font.Size = 12;
                    parrafoDes112.Font.Name = "Calibri (Cuerpo)";
                    parrafoDes112.Font.Bold = 0;

                   

                    string rutaTemporal = "";
                    using (MemoryStream ms = new MemoryStream(item.FirmaElectronica))
                    {
                        Image imagen = Image.FromStream(ms);
                        rutaTemporal = $@"{Environment.CurrentDirectory}\{item.SectorL}_{DateTime.Now.ToString("yyyyMMddhhmmss")}.png";
                        File.WriteAllBytes(rutaTemporal, ms.ToArray());
                    }

                    Word.Range parrafoFirma = parrafoDescripcion.Range;
                    Word.Shape myShape = document.Shapes.AddPicture(rutaTemporal, false, true, 0, 0, document.Application.CentimetersToPoints((float)21), document.Application.CentimetersToPoints((float)29.7), parrafoFirma);

                    // Eliminar la imagen temporal
                    if (File.Exists(rutaTemporal))
                    {
                        File.Delete(rutaTemporal);
                    }

                    // Ajustar el tamaño de la imagen y configurar sus propiedades
                    myShape.ScaleHeight(0.5f, MsoTriState.msoTrue);
                    myShape.ScaleWidth(0.5f, MsoTriState.msoTrue);

                    // Configurar el formato de ajuste de texto
                    myShape.WrapFormat.Type = WdWrapType.wdWrapTight;

                    // Centrar la imagen horizontalmente
                    myShape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    myShape.Left = (document.PageSetup.PageWidth - myShape.Width) / 2;

                    // Ajustar la posición vertical
                    float verticalPositionInCm = 19.01f; // Cambia este valor según la posición vertical deseada en cm
                    float verticalPositionInPoints = verticalPositionInCm * 28.35f; // Convertir cm a puntos
                    myShape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    myShape.Top = verticalPositionInPoints;

                    // Mantener el formato de ajuste de texto detrás del texto
                    myShape.WrapFormat.Type = WdWrapType.wdWrapBehind;
                    myShape.ZOrder(MsoZOrderCmd.msoSendBackward);

                    // Mostrar el documento
                    //wordApp.Visible = true;


                    // Agregar un nuevo párrafo al documento
                    Word.Paragraph parrafoFooter2 = document.Content.Paragraphs.Add(ref missing);
                    parrafoFooter2.Range.Text = $@"{Environment.NewLine}{Environment.NewLine}______________________________________________";
                    parrafoFooter2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    parrafoFooter2.Range.InsertParagraphAfter();

                    //NOMBRE LIDER
                    Word.Paragraph parrafoN = document.Content.Paragraphs.Add(ref missing);
                    parrafoN.Range.Text = $@"";
                    parrafoN.Range.Font.Size = 12;
                    parrafoN.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoN.Range.Font.Bold = 0;

                    Word.Range parrafoN1 = parrafoN.Range;
                    parrafoN1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoN1.Text = $@"{item.NombreL}";
                    parrafoN1.Font.Size = 12;
                    parrafoN1.Font.Name = "Calibri (Cuerpo)";
                    parrafoN1.Font.Bold = 1;

                    parrafoN.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    parrafoN.Range.InsertParagraphAfter();

                    // LIDER
                    Word.Paragraph parrafoLider = document.Content.Paragraphs.Add(ref missing);
                    parrafoLider.Range.Text = $@"Líder de Cobro Sector ";
                    parrafoLider.Range.Font.Size = 12;
                    parrafoLider.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoLider.Range.Font.Bold = 0;

                    Word.Range parrafoLider1 = parrafoLider.Range;
                    parrafoLider1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoLider1.Text = $@"{item.Area}";
                    parrafoLider1.Font.Size = 12;
                    parrafoLider1.Font.Name = "Calibri (Cuerpo)";
                    parrafoLider1.Font.Bold = 1;

                    Word.Paragraph parrafoSector = document.Content.Paragraphs.Add(ref missing);
                    parrafoSector.Range.Text = $@"{item.Email} ";
                    parrafoSector.Range.Font.Size = 12;
                    parrafoSector.Range.Font.Name = "Calibri (Cuerpo)";
                    parrafoSector.Range.Font.Bold = 0;

                    parrafoLider.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    parrafoLider.Range.InsertParagraphAfter();

                    Word.Range parrafoLider2 = parrafoLider.Range;
                    parrafoLider2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    parrafoLider2.Text = $@"{Environment.NewLine}Gestor de Cobro: {item.nombreg}";
                    parrafoLider2.Font.Size = 8;
                    parrafoLider2.Font.Name = "Calibri (Cuerpo)";
                    parrafoLider2.Font.Bold = 0;

                    parrafoLider.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    parrafoLider.Range.InsertParagraphAfter();

                    // Guardar y cerrar el documento por cada cliente en word
                    string filePath = $@"{directorioDestino}\Notificación de Deuda Prejurídico_Clave_{item.Clave}.docx";
                    document.SaveAs2(filePath);
                    document.Close();
                    ReleaseObject(document);
                    ReleaseObject(wordApp);
                    archivos.Add(filePath);

                }
                catch (Exception ex)
                {
                    return ex.ToString();
                }
            }

            GuardarComoPDF(archivos, directorioDestino);
            EliminarArchivos(archivos);
            //var result = UnirArchivos(archivos, directorioDestino);
            return "success";

        }
        static bool EliminarArchivos(List<string> archivos)
        {
            foreach (var archivo in archivos)
            {
                File.Delete(archivo);
            }
            return true;
        }

        static bool GuardarComoPDF(List<string> archivos, string directorioDestino)
        {
            List<string> archivosPDF = new();

            foreach (var archivo in archivos)
            {
                _Application oWord = null;
                _Document oDoc = null;

                try
                {
                    // Inicializar Word Interop
                    oWord = new Word.Application { Visible = false };

                    object oMissing = System.Reflection.Missing.Value;
                    object oInput = archivo;
                    object oOutput = archivo.Replace(".docx", ".pdf");
                    object oFormat = WdSaveFormat.wdFormatPDF;

                    // Abrir el documento
                    oDoc = oWord.Documents.Open(
                        ref oInput, ref oMissing, true, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, true, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    oDoc.Activate();

                    // Guardar como PDF
                    oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    archivosPDF.Add(archivo.Replace(".docx", ".pdf"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al convertir {archivo} a PDF: {ex.Message}");
                    return false;
                }
                finally
                {
                    // Liberar los objetos COM
                    if (oDoc != null)
                    {
                        oDoc.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                    }
                    if (oWord != null)
                    {
                        oWord.Quit(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

            return true;
        }


        static bool KillProcess(string processName)
        {
            try
            {
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo("taskkill.exe", $"/f /im {processName}");

                startInfo.UseShellExecute = false;  // Necesario para redirigir la salida
                startInfo.CreateNoWindow = true;    // Evita la creación de una ventana de consola
                startInfo.RedirectStandardOutput = true;

                process.StartInfo = startInfo;
                process.Start();

                string output = process.StandardOutput.ReadToEnd();
                Console.WriteLine(output);

                process.Close();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al intentar terminar {processName}: {ex.Message}");
                return false;
            }
        }

        static bool UnirArchivos(List<string> archivos, string directorioDestino)
        {
            PdfDocument pdfDestino = new PdfDocument();

            foreach (var archivo in archivos)
            {
                PdfDocument inputPDFDocument = PdfReader.Open(archivo, PdfDocumentOpenMode.Import);
                pdfDestino.Version = inputPDFDocument.Version;
                foreach (PdfPage page in inputPDFDocument.Pages)
                {
                    pdfDestino.AddPage(page);
                }

            }
            pdfDestino.Save($@"{directorioDestino}\Notificaciones {DateTime.Now.ToString("yyyyMMddHHmmss")}.pdf");

            return true;
        }

        static void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }

        static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Error al liberar objeto: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

    }

}
