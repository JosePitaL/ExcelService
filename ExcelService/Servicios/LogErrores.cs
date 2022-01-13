using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService.Servicios
{
    public static class LogErrores
    {
        public static void CrearLogErrores(List<Modelo.Error> errores, string PathNoAptos, string Fichero)
        {
            if(!File.Exists(PathNoAptos + Fichero.Replace(".xslm","") + ".txt"))
            {
                File.Create(PathNoAptos + Fichero.Replace(".xslm", "") + ".txt").Close();
            }

            foreach (var item in errores)
            {
                File.AppendAllText(PathNoAptos + Fichero.Replace(".xslm", "") + ".txt", "Hay un error en el registro " + item.Numero_Fila + " en la columna " + item.Numero_Columna + "\n");
            }
        }
    }
}
