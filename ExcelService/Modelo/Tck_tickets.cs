using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService.Modelo
{
    public class Tck_tickets
    {
        public int Id { get; set; }
        public int Numero_ticket { get; set; }
        public string Asunto { get; set; }
        public DateTime Fecha_creacion { get; set; }
        public string Nombre_plantilla { get; set; }
        public string Nombre_archivo { get; set; }
        public int Id_Estado { get; set; }
        public DateTime Fecha_registro { get; set; }
        public DateTime Fecha_modificacion { get; set; }
    }
}
