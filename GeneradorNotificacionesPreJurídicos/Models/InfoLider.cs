using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneradorNotificacionesPreJurídicos.Models
{
    public class InfoLider
    {
        public string? SectorL { get; set; }
        public string? NombreL { get; set; }

        public string? AnalistaC {  get; set; }

        public string? Telefono { get; set; }

        public string? Email { get; set; }

        public byte[]? FirmaElectronica { get; set; }
    }
}
