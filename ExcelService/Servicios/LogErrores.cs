using Dapper;
using MySql.Data.MySqlClient;
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
        public static void CrearLogErrores(List<Modelo.Error> errores, string PathNoAptos, string Fichero, int Ticket, int IdBaseDatos)
        {
            if(!File.Exists(PathNoAptos + Fichero.Replace(".xslm","") + ".txt"))
            {
                File.Create(PathNoAptos + Fichero.Replace(".xslm", "") + ".txt").Close();
            }

            foreach (var item in errores)
            {
                File.AppendAllText(PathNoAptos + Fichero.Replace(".xslm", "") + ".txt", "Hay un error en el registro " + item.Numero_Fila + " en la columna " + item.Numero_Columna + "\n");
                using (MySqlConnection conexion = new MySqlConnection("Server=localhost;Database=paydistrict; Uid=root;Pwd=020Na@es"))
                {
                    conexion.Execute("INSERT INTO tck_errores_validacion (id_ticket,columna,fila,fecha_registro) VALUES(" + IdBaseDatos + ",'"+item.Numero_Columna+"',"+item.Numero_Fila+",CURDATE())");
                }
            }
        }
    }
}
