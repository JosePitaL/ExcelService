using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using ExcelService.Modelo;
using ExcelService.Servicios;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace ExcelService
{
    class Program
    {
        
    
        static void Main(string[] args)
        {
            //ARGUMENTOS
            string PathAptos = @"C:\Users\34645\Desktop\PAYDISTRICT\APTO\";
            string PathNoAptos = @"C:\Users\34645\Desktop\PAYDISTRICT\NO APTO\";
            int Ticket = 37;
            string PathExcelEntrada = @"C:\Users\34645\Desktop\PAYDISTRICT\PLANTILLA PRUEBA OZONA.xlsm";
            string PathExcelSalida = @"C:\Users\34645\Desktop\PAYDISTRICT\Plantilla alta masiva - copia.xlsx";

            bool registro = true;
            bool Con_Errores = false;
            int fila = 3;
            int contadorTrabajadores = 1;
            Trabajador trabajador;
            List<Trabajador> trabajadores = new List<Trabajador>();

            var excel_entrada = new Application();
            excel_entrada.Visible = true;
            var pestañas_excel_entrada = excel_entrada.Workbooks.Open(PathExcelEntrada);

            while (registro)
            {
                if (excel_entrada.Cells[fila, 1].Value is null && excel_entrada.Cells[fila, 2].Value is null && excel_entrada.Cells[fila, 3].Value is null)
                {
                    registro = false;
                    continue;
                }
                else
                {

                    ValidacionEntrada.Validar(excel_entrada, fila, PathNoAptos, PathExcelEntrada.Split('\\').Last(), Ticket, 1,  out trabajador);
                    if (trabajador != null)
                    {
                        trabajadores.Add(trabajador);
                        
                    }
                    else
                    {
                        Con_Errores = true;
                    }

                }
                contadorTrabajadores++;
                fila++;
            }


            if(Con_Errores)
            {
                pestañas_excel_entrada.SaveAs(PathNoAptos + PathExcelEntrada.Split('\\').Last()); //SAVE AS  RENOMBRADO
                pestañas_excel_entrada.Close();
            }
            else
            {
                pestañas_excel_entrada.Close();
                SalidaExcel.CrearExcelSalida(trabajadores, PathExcelSalida, PathAptos, Ticket, contadorTrabajadores);
            }

            foreach (var proceso in Process.GetProcesses())
            {
                if (proceso.ProcessName.ToUpper().Contains("EXCEL"))
                    proceso.Kill();
            }  
        }
    }
}
