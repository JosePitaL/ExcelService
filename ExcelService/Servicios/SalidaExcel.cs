using ExcelService.Modelo;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService.Servicios
{
    public static class SalidaExcel
    {
        public static void CrearExcelSalida(List<Trabajador> trabajadores, string PathExcelSalida, string PathAptos, int Ticket)
        {
            var excel_salida = new Application();
            excel_salida.Visible = true;
            var pestañas_excel_salida = excel_salida.Workbooks.Open(PathExcelSalida);

            int fila = 3;
            string codigoempresa = trabajadores[0].Codigo_Empresa.ToString();

            foreach (var trabajador in trabajadores)
            {
                excel_salida.Cells[fila, 1].Value = trabajador.Codigo_Empresa;
                excel_salida.Cells[fila, 2].Value = trabajador.Codigo_Centro;
                excel_salida.Cells[fila, 3].Value = trabajador.Codigo_Trabajador;
                excel_salida.Cells[fila, 4].Value = trabajador.Tipo_Documento;
                excel_salida.Cells[fila, 5].Value = trabajador.Documento;
                excel_salida.Cells[fila, 6].Value = trabajador.Nombre;
                excel_salida.Cells[fila, 7].Value = trabajador.Primer_Apellido;
                if(!trabajador.Segundo_Apellido.Equals(""))
                {
                    excel_salida.Cells[fila, 8].Value = trabajador.Segundo_Apellido;
                }
                excel_salida.Cells[fila, 9].Value = trabajador.Sexo;
                excel_salida.Cells[fila, 10].Value = trabajador.Naf;
                excel_salida.Cells[fila, 11].Value = trabajador.Fecha_Alta;
                excel_salida.Cells[fila, 12].Value = trabajador.Fecha_Nacimiento;
                excel_salida.Cells[fila, 13].Value = trabajador.Nacionalidad;
                excel_salida.Cells[fila, 14].Value = trabajador.Email_Profesional;
                excel_salida.Cells[fila, 15].Value = trabajador.Codigo_Convenio;
                excel_salida.Cells[fila, 16].Value = trabajador.Codigo_Categoria;
                excel_salida.Cells[fila, 17].Value = trabajador.Codigo_Puesto;
                if(!trabajador.Grupo_Antiguedad.Equals(""))
                {
                    excel_salida.Cells[fila, 18].Value = trabajador.Grupo_Antiguedad;
                }
                excel_salida.Cells[fila, 19].Value = trabajador.Grupo_Pagas_Extra;
                if (!trabajador.Grupo_Complemento_It.Equals(""))
                {
                    excel_salida.Cells[fila, 20].Value = trabajador.Grupo_Complemento_It;
                }
                excel_salida.Cells[fila, 21].Value = trabajador.Regimen;
                excel_salida.Cells[fila, 22].Value = trabajador.Grupo_Tarifa;
                excel_salida.Cells[fila, 23].Value = trabajador.Tipo_Cobro;
                if (!trabajador.Ocupacion_Tgss.Equals(""))
                {
                    excel_salida.Cells[fila, 24].Value = trabajador.Ocupacion_Tgss;
                }
                excel_salida.Cells[fila, 25].Value = trabajador.Entidad;
                excel_salida.Cells[fila, 26].Value = trabajador.Agencia;
                excel_salida.Cells[fila, 27].Value = trabajador.Dc;
                excel_salida.Cells[fila, 28].Value = trabajador.Cuenta;
                excel_salida.Cells[fila, 29].Value = trabajador.Iban;
                if (!trabajador.Swift_Bic.Equals(""))
                {
                    excel_salida.Cells[fila, 30].Value = trabajador.Swift_Bic;
                }
                excel_salida.Cells[fila, 31].Value = trabajador.Tipo_Contrato;
                excel_salida.Cells[fila, 32].Value = trabajador.Tipo_Cotizacion;
                excel_salida.Cells[fila, 33].Value = trabajador.Tipo_Bruto_Anual;
                excel_salida.Cells[fila, 34].Value = trabajador.Bruto_Anual;
                excel_salida.Cells[fila, 35].Value = trabajador.Cno_Ocupacion;
                excel_salida.Cells[fila, 36].Value = trabajador.Nivel_Formativo;
                excel_salida.Cells[fila, 37].Value = trabajador.Fecha_Inicio_Contrato;
                if (trabajador.Meses != -1)
                {
                    excel_salida.Cells[fila, 38].Value = trabajador.Meses;
                }
                if (trabajador.Dias != -1)
                {
                    excel_salida.Cells[fila, 39].Value = trabajador.Dias;
                }
                if (trabajador.Fecha_Fin_Contrato.Year != 1900)
                {
                    excel_salida.Cells[fila, 40].Value = trabajador.Fecha_Fin_Contrato;
                }
                excel_salida.Cells[fila, 41].Value = trabajador.Fecha_Pagas_Extra;
                excel_salida.Cells[fila, 42].Value = trabajador.Fecha_Antiguedad;
                excel_salida.Cells[fila, 43].Value = trabajador.Fecha_Antiguedad_Empresa;
                fila++;
            }

            fila = 3;

            foreach (var trabajador in trabajadores)
            {
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 1].Value = trabajador.Codigo_Empresa;
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 2].Value = trabajador.Codigo_Trabajador;
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 3].Value = trabajador.Nombre;
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 4].Value = trabajador.Tipo_Via;
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 5].Value = trabajador.Via_Publica;
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 6].Value = trabajador.Numero;
                if (!trabajador.Escalera.Equals(""))
                {
                    excel_salida.Sheets["Datos Domicilio"].Cells[fila, 7].Value = trabajador.Escalera;
                }
                if (trabajador.Piso != -1)
                {
                    excel_salida.Sheets["Datos Domicilio"].Cells[fila, 8].Value = trabajador.Piso;
                }
                if (trabajador.Puerta != -1)
                {
                    excel_salida.Sheets["Datos Domicilio"].Cells[fila, 9].Value = trabajador.Puerta;
                }
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 10].Value = trabajador.Pais;
                excel_salida.Sheets["Datos Domicilio"].Cells[fila, 11].Value = trabajador.Codigo_Postal;
                fila++;
            }

            fila = 3;

            foreach (var trabajador in trabajadores)
            {
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 1].Value = trabajador.Codigo_Empresa;
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 2].Value = trabajador.Codigo_Trabajador;
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 3].Value = trabajador.Nombre;
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 4].Value = "NO";
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 7].Value = "Automático";
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 9].Value = trabajador.Indicador_No_Residente;
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 10].Value = trabajador.Clave_Percepcion;
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 11].Value = trabajador.Situacion_Familiar;
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 12].Value = trabajador.Documento_Conyugue;
                if (trabajador.Discapacidad != -1)
                {
                    excel_salida.Sheets["Datos IRPF"].Cells[fila, 13].Value = trabajador.Discapacidad;
                }
                excel_salida.Sheets["Datos IRPF"].Cells[fila, 14].Value = trabajador.Con_Ayuda;
                fila++;
            }

            fila = 3;
            int numero_empleado = 1;

            foreach (var trabajador in trabajadores)
            {
                numero_empleado = 1;
                foreach (var descendiente in trabajador.Descendientes)
                {
                    excel_salida.Sheets["Datos IRPF (Descendientes)"].Cells[fila, 1].Value = trabajador.Codigo_Empresa;
                    excel_salida.Sheets["Datos IRPF (Descendientes)"].Cells[fila, 2].Value = trabajador.Codigo_Trabajador;
                    excel_salida.Sheets["Datos IRPF (Descendientes)"].Cells[fila, 3].Value = trabajador.Nombre;
                    excel_salida.Sheets["Datos IRPF (Descendientes)"].Cells[fila, 4].Value = "Nuevo";
                    excel_salida.Sheets["Datos IRPF (Descendientes)"].Cells[fila, 5].Value = numero_empleado;
                    excel_salida.Sheets["Datos IRPF (Descendientes)"].Cells[fila, 6].Value = descendiente.Ano_Nacimiento;
                    excel_salida.Sheets["Datos IRPF (Descendientes)"].Cells[fila, 7].Value = "50%";
                    numero_empleado++;
                    fila++;
                }
            }

            fila = 3;

            foreach (var trabajador in trabajadores)
            {
                numero_empleado = 1;
                foreach (var imputacion in trabajador.Imputaciones)
                {
                    excel_salida.Sheets["Imputación"].Cells[fila, 1].Value = trabajador.Codigo_Empresa;
                    excel_salida.Sheets["Imputación"].Cells[fila, 2].Value = trabajador.Codigo_Trabajador;
                    excel_salida.Sheets["Imputación"].Cells[fila, 3].Value = trabajador.Nombre;
                    excel_salida.Sheets["Imputación"].Cells[fila, 4].Value = "NO";
                    excel_salida.Sheets["Imputación"].Cells[fila, 5].Value = imputacion._Imputacion;
                    excel_salida.Sheets["Imputación"].Cells[fila, 6].Value = imputacion.Porcentaje;
                    fila++;
                }
            }
            pestañas_excel_salida.SaveAs(PathAptos + codigoempresa + "_" + Ticket + "_" + DateTime.Now.ToString("ddMMyyyyHHmm") + "_" + trabajadores.Count);
            pestañas_excel_salida.Close();
        }       
    }
}
