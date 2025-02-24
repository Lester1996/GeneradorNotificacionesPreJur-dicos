using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GeneradorNotificacionesPreJurídicos.Models
{
    public class InformacionCliente
    {
        public string Clave { get; set; }
        public int NisRad { get; set; }
        public string? NombreCliente { get; set; }
        public string? DireccionCliente { get; set; }
        public string? Area {  get; set; }
        public DateTime?  FechaUltimoPago { get; set; }
        public DateTime? FechaActual {  get; set; }

        public int DiaActual { get; set; }
        public int MesActual { get; set; }
        public int AnioActual { get; set; }
        
        public string? MesActualText {  get; set; }
        public string? DiaActualText { get; set; }
        public double DeudaTotal { get; set; }



    }
}


