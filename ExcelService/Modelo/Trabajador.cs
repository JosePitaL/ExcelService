using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService.Modelo
{
    public class Trabajador
    {
        
        public double? Codigo_Empresa { get; set; }
        public double? Codigo_Centro { get; set; }
        public double? Codigo_Trabajador { get; set; }
        public string Tipo_Documento { get; set; }
        public string Documento { get; set; }
        public string Nombre { get; set; }
        public string Primer_Apellido { get; set; }
        public string Segundo_Apellido { get; set; }
        public string Sexo { get; set; }
        public string Naf { get; set; }
        public DateTime? Fecha_Alta { get; set; }
        public DateTime? Fecha_Nacimiento { get; set; }
        public string Nacionalidad { get; set; }
        public string Email_Profesional { get; set; }
        public string Codigo_Convenio { get; set; }
        public string Codigo_Categoria { get; set; }
        public double? Codigo_Puesto { get; set; }
        public string Grupo_Antiguedad { get; set; }
        public string Grupo_Pagas_Extra { get; set; }
        public string Grupo_Complemento_It { get; set; }
        public string Regimen { get; set; }
        public string Grupo_Tarifa { get; set; }
        public string Tipo_Cobro { get; set; }
        public string Ocupacion_Tgss { get; set; }
        public string Entidad { get; set; }
        public string Agencia { get; set; }
        public string Dc { get; set; }
        public string Cuenta { get; set; }
        public string Iban { get; set; }
        public string Swift_Bic { get; set; }
        public string Tipo_Contrato { get; set; }
        public string Tipo_Cotizacion { get; set; }
        public string Tipo_Bruto_Anual { get; set; }
        public double? Bruto_Anual { get; set; }
        public string Cno_Ocupacion { get; set; }
        public string Nivel_Formativo { get; set; }
        public DateTime? Fecha_Inicio_Contrato { get; set; }
        public double? Meses { get; set; }
        public double? Dias { get; set; }
        public DateTime Fecha_Fin_Contrato { get; set; }
        public DateTime? Fecha_Pagas_Extra { get; set; }
        public DateTime? Fecha_Antiguedad { get; set; }
        public DateTime? Fecha_Antiguedad_Empresa { get; set; }
        public string Tipo_Via { get; set; }
        public string Via_Publica { get; set; }
        public string Numero { get; set; }
        public string Escalera { get; set; }
        public string Piso { get; set; }
        public string Puerta { get; set; }
        public string Pais { get; set; }
        public string Codigo_Postal { get; set; }
        public string Indicador_No_Residente { get; set; }
        public string Clave_Percepcion { get; set; }
        public string Situacion_Familiar { get; set; }
        public string Documento_Conyugue { get; set; }
        public double? Discapacidad { get; set; }
        public string Con_Ayuda { get; set; }
        public List<Descendiente> Descendientes { get; set; }
        public List<Imputacion> Imputaciones { get; set; }

    }
}
