using ExcelService.Modelo;
using ExcelService.Servicios;
using Microsoft.Office.Interop.Excel;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    public class ExcelService : CodeActivity
    {
        [Category("Output")]
        public OutArgument<bool> _ResultadoFichero { get; set; }
        [Category("Output")]
        public OutArgument<string> _NombreExcelResultado { get; set; }

        [Category("Input")]
        public InArgument<string> _PathAptos { get; set; }
        [Category("Input")]
        public InArgument<int> _IdBaseDatos { get; set; }
        [Category("Input")]
        public InArgument<string> _PathNoAptos { get; set; }

        [Category("Input")]
        public InArgument<string> _PathExcelEntrada { get; set; }
        [Category("Input")]
        public InArgument<string> _PathExcelSalida { get; set; }
        [Category("Input")]
        public InArgument<int> _Ticket { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            //ARGUMENTOS
            string PathAptos = _PathAptos.Get(context);
            string PathNoAptos = _PathNoAptos.Get(context);
            int Ticket = _Ticket.Get(context);
            int IdBaseDatos = _IdBaseDatos.Get(context);
            string PathExcelEntrada = _PathExcelEntrada.Get(context);
            string PathExcelSalida = _PathExcelSalida.Get(context);

            bool registro = true;
            bool Con_Errores = false;
            int contadorTrabajadores = 1;
            int fila = 3;
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

                    ValidacionEntrada.Validar(excel_entrada, fila, PathNoAptos, PathExcelEntrada.Split('\\').Last(), Ticket, IdBaseDatos, out trabajador);
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


            if (Con_Errores)
            {
                pestañas_excel_entrada.SaveAs(PathNoAptos + PathExcelEntrada.Split('\\').Last()); //SAVE AS  RENOMBRADO
                pestañas_excel_entrada.Close();
                _ResultadoFichero.Set(context, false);
            }
            else
            {
                pestañas_excel_entrada.Close();
                _NombreExcelResultado.Set(context,SalidaExcel.CrearExcelSalida(trabajadores, PathExcelSalida, PathAptos, Ticket, contadorTrabajadores));
                _ResultadoFichero.Set(context, true);
            }

            foreach (var proceso in Process.GetProcesses())
            {
                if (proceso.ProcessName.ToUpper().Contains("EXCEL"))
                    proceso.Kill();
            }
        }
    }
}
