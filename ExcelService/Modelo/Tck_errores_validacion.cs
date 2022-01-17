using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService.Modelo
{
    public class Tck_errores_validacion
    {
        public int Id { get; set; }
        public string Columna { get; set; }
        public int Fila { get; set; }
        public string Error { get; set; }
        public DateTime Fecha_registro { get; set; }
    }
}
