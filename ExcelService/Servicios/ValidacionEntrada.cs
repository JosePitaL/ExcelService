using ExcelService.Modelo;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService.Servicios
{
    public static class ValidacionEntrada
    {
        public static void Validar(Application excel_entrada, int fila, string PathNoAptos, string Fichero, int Ticket, int IdBaseDatos, out Trabajador trabajador)
        {
            List<Modelo.Error> errores = new List<Modelo.Error>();
            bool Con_Errores = false;
            trabajador = new Trabajador();

            //Codigo_Empresa
            if (excel_entrada.Cells[fila, 1].Value is null)
            {
                excel_entrada.Cells[fila, 1].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 1].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Codigo_Empresa = excel_entrada.Cells[fila, 1].Value;
            }

            //Codigo_Centro
            if (excel_entrada.Cells[fila, 2].Value is null)
            {
                excel_entrada.Cells[fila, 2].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 2].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Codigo_Centro = excel_entrada.Cells[fila, 2].Value;
            }

            //Codigo_Trabajador
            if (excel_entrada.Cells[fila, 3].Value is null)
            {
                excel_entrada.Cells[fila, 3].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 3].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Codigo_Trabajador = excel_entrada.Cells[fila, 3].Value;
            }

            //Tipo_Documento
            if (excel_entrada.Cells[fila, 4].Value is null)
            {
                excel_entrada.Cells[fila, 4].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 4].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Tipo_Documento = excel_entrada.Cells[fila, 4].Value.ToString();
            }

            //Documento
            if (excel_entrada.Cells[fila, 5].Value is null)
            {
                excel_entrada.Cells[fila, 5].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 5].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Documento = excel_entrada.Cells[fila, 5].Value.ToString();
            }

            //Nombre
            if (excel_entrada.Cells[fila, 6].Value is null)
            {
                excel_entrada.Cells[fila, 6].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 6].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Nombre = excel_entrada.Cells[fila, 6].Value.ToString();
            }

            //Primer_Apellido
            if (excel_entrada.Cells[fila, 7].Value is null)
            {
                excel_entrada.Cells[fila, 7].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 7].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Primer_Apellido = excel_entrada.Cells[fila, 7].Value.ToString();
            }

            //Segundo_Apellido
            if (excel_entrada.Cells[fila, 8].Value is null)
            {
                trabajador.Segundo_Apellido = "";
            }
            else
            {
                trabajador.Segundo_Apellido = excel_entrada.Cells[fila, 8].Value.ToString();
            }

            //Sexo
            if (excel_entrada.Cells[fila, 9].Value is null)
            {
                excel_entrada.Cells[fila, 9].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 9].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Sexo = excel_entrada.Cells[fila, 9].Value.ToString();
            }

            //Naf
            if (excel_entrada.Cells[fila, 10].Value is null)
            {
                excel_entrada.Cells[fila, 10].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 10].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Naf = excel_entrada.Cells[fila, 10].Value.ToString();
            }

            //Fecha_Alta
            if (excel_entrada.Cells[fila, 11].Value is null)
            {
                excel_entrada.Cells[fila, 11].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 11].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Fecha_Alta = excel_entrada.Cells[fila, 11].Value;
            }

            //Fecha_Nacimiento
            if (excel_entrada.Cells[fila, 12].Value is null)
            {
                excel_entrada.Cells[fila, 12].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 12].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Fecha_Nacimiento = excel_entrada.Cells[fila, 12].Value;
            }

            //Nacionalidad
            if (excel_entrada.Cells[fila, 13].Value is null)
            {
                excel_entrada.Cells[fila, 13].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 13].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Nacionalidad = excel_entrada.Cells[fila, 13].Value.ToString();
            }

            //Email_Profesional
            if (excel_entrada.Cells[fila, 14].Value is null)
            {
                excel_entrada.Cells[fila, 14].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 14].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Email_Profesional = excel_entrada.Cells[fila, 14].Value.ToString();
            }

            //Codigo_Convenio
            if (excel_entrada.Cells[fila, 15].Value is null)
            {
                excel_entrada.Cells[fila, 15].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 15].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Codigo_Convenio = excel_entrada.Cells[fila, 15].Value.ToString();
            }

            //Codigo_Categoria
            if (excel_entrada.Cells[fila, 16].Value is null)
            {
                excel_entrada.Cells[fila, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 16].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Codigo_Categoria = excel_entrada.Cells[fila, 16].Value.ToString();
            }

            //Codigo_Puesto
            if (excel_entrada.Cells[fila, 17].Value is null)
            {
                excel_entrada.Cells[fila, 17].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 17].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Codigo_Puesto = double.Parse(excel_entrada.Cells[fila, 17].Value.ToString());
            }

            //Grupo_Antiguedad
            if (excel_entrada.Cells[fila, 18].Value is null)
            {
                trabajador.Grupo_Antiguedad = "";
            }
            else
            {
                trabajador.Grupo_Antiguedad = excel_entrada.Cells[fila, 18].Value.ToString();
            }

            //Grupo_Pagas_Extra
            if (excel_entrada.Cells[fila, 19].Value is null)
            {
                trabajador.Grupo_Pagas_Extra = "Pagas extras prorrateadas";
            }
            else
            {
                trabajador.Grupo_Pagas_Extra = excel_entrada.Cells[fila, 19].Value.ToString();
            }

            //Grupo_Complemento_It
            if (excel_entrada.Cells[fila, 20].Value is null)
            {
                trabajador.Grupo_Complemento_It = "";
            }
            else
            {
                trabajador.Grupo_Complemento_It = excel_entrada.Cells[fila, 20].Value.ToString();
            }

            //Regimen
            if (excel_entrada.Cells[fila, 21].Value is null)
            {
                trabajador.Regimen = "Régimen General";
            }
            else
            {
                trabajador.Regimen = excel_entrada.Cells[fila, 21].Value.ToString();
            }

            //Grupo_Tarifa
            if (excel_entrada.Cells[fila, 22].Value is null)
            {
                excel_entrada.Cells[fila, 22].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 22].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Grupo_Tarifa = excel_entrada.Cells[fila, 22].Value.ToString();
            }

            //Tipo_Cobro
            if (excel_entrada.Cells[fila, 23].Value is null)
            {
                trabajador.Tipo_Cobro = "Mensual";
            }
            else
            {
                trabajador.Tipo_Cobro = excel_entrada.Cells[fila, 23].Value.ToString();
            }

            //Ocupacion_Tgss
            if (excel_entrada.Cells[fila, 24].Value is null)
            {
                trabajador.Ocupacion_Tgss = "";
            }
            else
            {
                trabajador.Ocupacion_Tgss = excel_entrada.Cells[fila, 24].Value.ToString();
            }

            //Entidad
            if (excel_entrada.Cells[fila, 25].Value is null)
            {
                excel_entrada.Cells[fila, 25].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 25].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Entidad = excel_entrada.Cells[fila, 25].Value.ToString();
            }

            //Agencia
            if (excel_entrada.Cells[fila, 26].Value is null)
            {
                excel_entrada.Cells[fila, 26].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 26].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Agencia = excel_entrada.Cells[fila, 26].Value.ToString();
            }

            //Dc
            if (excel_entrada.Cells[fila, 27].Value is null)
            {
                excel_entrada.Cells[fila, 27].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 27].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Dc = excel_entrada.Cells[fila, 27].Value.ToString();
            }

            //Cuenta
            if (excel_entrada.Cells[fila, 28].Value is null)
            {
                excel_entrada.Cells[fila, 28].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 28].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Cuenta = excel_entrada.Cells[fila, 28].Value.ToString();
            }

            //Iban
            if (excel_entrada.Cells[fila, 29].Value is null)
            {
                excel_entrada.Cells[fila, 29].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 29].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Iban = excel_entrada.Cells[fila, 29].Value.ToString();
            }

            //Swift_Bic
            if (excel_entrada.Cells[fila, 30].Value is null)
            {
                trabajador.Swift_Bic = "";
            }
            else
            {
                trabajador.Swift_Bic = excel_entrada.Cells[fila, 30].Value.ToString();
            }

            //Tipo_Contrato
            if (excel_entrada.Cells[fila, 31].Value is null)
            {
                excel_entrada.Cells[fila, 31].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 31].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Tipo_Contrato = excel_entrada.Cells[fila, 31].Value.ToString();
            }

            //Tipo_Cotizacion
            if (excel_entrada.Cells[fila, 32].Value is null)
            {
                excel_entrada.Cells[fila, 32].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 32].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Tipo_Cotizacion = excel_entrada.Cells[fila, 32].Value.ToString();
            }

            //Tipo_Bruto_Anual
            if (excel_entrada.Cells[fila, 33].Value is null)
            {
                trabajador.Tipo_Bruto_Anual = "Según importe";
            }
            else
            {
                trabajador.Tipo_Bruto_Anual = excel_entrada.Cells[fila, 33].Value.ToString();
            }

            //Bruto_Anual
            if (excel_entrada.Cells[fila, 34].Value is null)
            {
                excel_entrada.Cells[fila, 34].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 34].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Bruto_Anual = double.Parse(excel_entrada.Cells[fila, 34].Value.ToString());
            }

            //Cno_Ocupacion
            if (excel_entrada.Cells[fila, 35].Value is null)
            {
                excel_entrada.Cells[fila, 35].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 35].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Cno_Ocupacion = excel_entrada.Cells[fila, 35].Value.ToString();
            }

            //Nivel_Formativo
            if (excel_entrada.Cells[fila, 36].Value is null)
            {
                excel_entrada.Cells[fila, 36].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 36].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Nivel_Formativo = excel_entrada.Cells[fila, 36].Value.ToString();
            }

            //Fecha_Inicio_Contrato
            if (excel_entrada.Cells[fila, 37].Value is null)
            {
                excel_entrada.Cells[fila, 37].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 37].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Fecha_Inicio_Contrato = excel_entrada.Cells[fila, 37].Value;
            }

            //Meses
            if (excel_entrada.Cells[fila, 38].Value is null)
            {
                trabajador.Meses = -1;
            }
            else
            {
                trabajador.Meses = double.Parse(excel_entrada.Cells[fila, 38].Value);
            }

            //Dias
            if (excel_entrada.Cells[fila, 39].Value is null)
            {
                trabajador.Dias = -1;
            }
            else
            {
                trabajador.Dias = double.Parse(excel_entrada.Cells[fila, 39].Value);
            }

            //Fecha_Fin_Contrato
            if (excel_entrada.Cells[fila, 40].Value is null)
            {
                trabajador.Fecha_Fin_Contrato = new DateTime(1900, 1, 1);
            }
            else
            {
                trabajador.Fecha_Fin_Contrato = excel_entrada.Cells[fila, 40].Value;
            }

            //Fecha_Pagas_Extra
            if (excel_entrada.Cells[fila, 41].Value.Equals(""))
            {
                excel_entrada.Cells[fila, 41].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 41].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Fecha_Pagas_Extra = excel_entrada.Cells[fila, 41].Value;
            }

            //Fecha_Antiguedad
            if (excel_entrada.Cells[fila, 42].Value.Equals(""))
            {
                excel_entrada.Cells[fila, 42].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 42].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Fecha_Antiguedad = excel_entrada.Cells[fila, 42].Value;
            }

            //Fecha_Antiguedad_Empresa 
            if (excel_entrada.Cells[fila, 43].Value.Equals(""))
            {
                excel_entrada.Cells[fila, 43].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 43].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Fecha_Antiguedad_Empresa = excel_entrada.Cells[fila, 43].Value;
            }

            //Tipo_Via
            if (excel_entrada.Cells[fila, 44].Value is null)
            {
                excel_entrada.Cells[fila, 44].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 44].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Tipo_Via = excel_entrada.Cells[fila, 44].Value.ToString();
            }

            //Via_Publica
            if (excel_entrada.Cells[fila, 45].Value is null)
            {
                excel_entrada.Cells[fila, 45].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 45].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Via_Publica = excel_entrada.Cells[fila, 45].Value.ToString();
            }

            //Numero
            if (excel_entrada.Cells[fila, 46].Value is null)
            {
                excel_entrada.Cells[fila, 46].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 46].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Numero = excel_entrada.Cells[fila, 46].Value.ToString();
            }

            //Escalera
            if (excel_entrada.Cells[fila, 47].Value is null)
            {
                trabajador.Escalera = "";
            }
            else
            {
                trabajador.Escalera = excel_entrada.Cells[fila, 47].Value.ToString();
            }

            //Piso
            if (excel_entrada.Cells[fila, 48].Value is null)
            {
                trabajador.Piso = "";
            }
            else
            {
                trabajador.Piso = excel_entrada.Cells[fila, 48].Value.ToString();
            }

            //Puerta
            if (excel_entrada.Cells[fila, 49].Value is null)
            {
                trabajador.Puerta = "";
            }
            else
            {
                trabajador.Puerta = excel_entrada.Cells[fila, 49].Value.ToString();
            }

            //Pais
            if (excel_entrada.Cells[fila, 50].Value is null)
            {
                excel_entrada.Cells[fila, 50].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 50].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Pais = excel_entrada.Cells[fila, 50].Value.ToString();
            }

            //Codigo_Postal
            if (excel_entrada.Cells[fila, 51].Value is null)
            {
                excel_entrada.Cells[fila, 51].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 51].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Codigo_Postal = excel_entrada.Cells[fila, 51].Value.ToString();
            }

            //Indicador_No_Residente
            if (excel_entrada.Cells[fila, 52].Value is null)
            {
                excel_entrada.Cells[fila, 52].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 52].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Indicador_No_Residente = excel_entrada.Cells[fila, 52].Value.ToString();
            }

            //Clave_Percepcion
            if (excel_entrada.Cells[fila, 53].Value is null)
            {
                excel_entrada.Cells[fila, 53].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 53].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Clave_Percepcion = excel_entrada.Cells[fila, 53].Value.ToString();
            }

            //Situacion_Familiar
            if (excel_entrada.Cells[fila, 54].Value is null)
            {
                excel_entrada.Cells[fila, 54].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                errores.Add(new Modelo.Error()
                {
                    Numero_Fila = fila,
                    Numero_Columna = excel_entrada.Cells[2, 54].Value
                });
                Con_Errores = true;
            }
            else
            {
                trabajador.Situacion_Familiar = excel_entrada.Cells[fila, 54].Value.ToString();
            }

            //Documento_Conyugue
            if (excel_entrada.Cells[fila, 55].Value is null && excel_entrada.Cells[fila, 54].Value != null)
            {
                if(excel_entrada.Cells[fila, 54].Value.Contains("2"))
                {
                    excel_entrada.Cells[fila, 55].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                    errores.Add(new Modelo.Error()
                    {
                        Numero_Fila = fila,
                        Numero_Columna = excel_entrada.Cells[2, 55].Value
                    });
                    Con_Errores = true;
                }
            }
            else
            {
                trabajador.Documento_Conyugue = excel_entrada.Cells[fila, 5].Value.ToString();
            }

            //Discapacidad
            if (excel_entrada.Cells[fila, 56].Value is null)
            {
                trabajador.Discapacidad = -1;
            }
            else
            {
                trabajador.Discapacidad = excel_entrada.Cells[fila, 56].Value;
            }

            //Con_Ayuda
            if (excel_entrada.Cells[fila, 57].Value is null)
            {
               trabajador.Con_Ayuda = "";
            }
            else
            {
                trabajador.Con_Ayuda = excel_entrada.Cells[fila, 57].Value.ToString();
            }

            //Descendientes
            trabajador.Descendientes = new List<Descendiente>();
            for (int i = 58; i <= 61; i++)
            {
                if(excel_entrada.Cells[fila, i].Value != null)
                {
                    trabajador.Descendientes.Add(new Descendiente()
                    {
                        Ano_Nacimiento = excel_entrada.Cells[fila, i].Value
                    });
                }
            }

            //Imputacion
            trabajador.Imputaciones = new List<Imputacion>();
            for (int i = 61; i <= 67; i = i + 2)
            {
                if (excel_entrada.Cells[fila, i].Value != null && excel_entrada.Cells[fila, i + 1].Value is null)
                {
                    excel_entrada.Cells[fila, i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                    excel_entrada.Cells[fila, i+1].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                    errores.Add(new Modelo.Error()
                    {
                        Numero_Fila = fila,
                        Numero_Columna = excel_entrada.Cells[2, i].Value
                    });
                    Con_Errores = true;
                }
                else if(excel_entrada.Cells[fila, i].Value != null && excel_entrada.Cells[fila, i + 1].Value != null)
                {
                    trabajador.Imputaciones.Add(new Imputacion()
                    {
                        _Imputacion = excel_entrada.Cells[fila, i].Value.ToString(),
                        Porcentaje = excel_entrada.Cells[fila, i + 1].Value,
                    });
                }
            }

            if(Con_Errores)
            {
                trabajador = null;
                LogErrores.CrearLogErrores(errores, PathNoAptos, Fichero, Ticket, IdBaseDatos);
            }
        }
    }
}
