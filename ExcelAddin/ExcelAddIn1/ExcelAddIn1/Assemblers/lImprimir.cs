using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAddIn1;
using ExcelAddIn.Access;

namespace ExcelAddIn1
{
    public class lImprimir
    {
        public string[,] _HojasSPR = new string[,] {
            {"Contribuyente".ToUpper()          , "31"  ,"3"    ,""                     ,"Contribuyente"},
            {"Contador".ToUpper()               , "35"  ,"3"    ,""                     ,"Contador"},
            {"Representante".ToUpper()          , "36"  ,"3"    ,""                     ,"Representante"},
            {"Generales".ToUpper()              , "446" ,"3"    ,""                     ,"Generales"},
            {"Anexo 1".ToUpper()                , "0"   ,"10"   ,""                     ,"1.- ESTADO DE SITUACION FINANCIERA"},
            {"Anexo 2".ToUpper()                , "0"   ,"9"    ,""                     ,"2.- ESTADO DE RESULTADO INTEGRAL"},
            {"Anexo 3".ToUpper()                , "0"   ,"22"   ,""                     ,"3.- ESTADO DE CAMBIOS EN EL CAPITAL CONTABLE"},
            {"Anexo 4".ToUpper()                , "0"   ,"5"    ,""                     ,"4.- ESTADO DE FLUJOS DE EFECTIVO "},
            {"Anexo 5".ToUpper()                , "0"   ,"14"   ,""                     ,"5.- INTEGRACION ANALITICA DE VENTAS O INGRESOS NETOS "},
            {"Anexo 6".ToUpper()                , "0"   ,"5"    ,""                     ,"6.- DETERMINACION DEL COSTO DE LO VENDIDO PARA EFECTOS CONTABLES Y DEL IMPUESTO SOBRE LA RENTA "},
            {"Anexo 7".ToUpper()                , "0"   ,"37"   ,""                     ,"7.- ANALISIS COMPARATIVO DE LAS SUBCUENTAS DE GASTOS"},
            {"Anexo 8".ToUpper()                , "0"   ,"9"    ,""                     ,"8.- ANALISIS COMPARATIVO DE LAS SUBCUENTAS DEL RESULTADO INTEGRAL DE FINANCIAMIENTO"},
            {"Anexo 9".ToUpper()                , "0"   ,"9"    ,""                     ,"9.- RELACION DE CONTRIBUCIONES A CARGO DEL CONTRIBUYENTE COMO SUJETO DIRECTO O EN SU CARACTER DE RETENEDOR"},
            {"Anexo 10".ToUpper()               , "0"   ,"15"   ,""                     ,"10.- RELACION DE CONTRIBUCIONES POR PAGAR"},
            {"Anexo 11".ToUpper()               , "0"   ,"4"    ,""                     ,"11.- CONCILIACION ENTRE EL RESULTADO CONTABLE Y FISCAL PARA EFECTOS DEL IMPUESTO SOBRE LA RENTA"},
            {"Anexo 12".ToUpper()               , "0"   ,"13"   ,"Generales|C96"        ,"12.- OPERACIONES FINANCIERAS DERIVADAS CONTRATADAS CON RESIDENTES EN EL EXTRANJERO "},
            {"Anexo 13".ToUpper()               , "0"   ,"10"   ,"Generales|C97"        ,"13.- INVERSIONES PERMANENTES EN SUBSIDIARIAS, ASOCIADAS Y AFILIADAS RESIDENTES EN EL EXTRANJERO"},
            {"Anexo 14".ToUpper()               , "0"   ,"12"   ,""                     ,"14.- SOCIOS O ACCIONISTAS QUE TUVIERON ACCIONES O PARTES SOCIALES"},
            {"Anexo 15".ToUpper()               , "0"   ,"4"    ,""                     ,"15.- CONCILIACION ENTRE LOS INGRESOS DICTAMINADOS SEGUN ESTADO DE RESULTADO INTEGRAL Y LOS ACUMULABLES PARA EFECTOS DEL IMPUESTO SOBRE LA RENTA Y  EL TOTAL DE ACTOS O ACTIVIDADES PARA EFECTOS DEL IMPUESTO AL VALOR AGREGADO"},
            {"Anexo 16".ToUpper()               , "0"   ,"11"   ,"Generales|C57"        ,"16.- OPERACIONES CON PARTES RELACIONADAS"},
            {"Anexo 17".ToUpper()               , "0"   ,"4"    ,"Generales|C57"        ,"17.- INFORMACION DEL CONTRIBUYENTE SOBRE SUS OPERACIONES CON PARTES RELACIONADAS"},
            {"Anexo 18".ToUpper()               , "0"   ,"4"    ,""                     ,"18.- DATOS INFORMATIVOS "},
            {"Anexo 19".ToUpper()               , "0"   ,"7"    ,"Generales|C98"        ,"19.- INFORMACION DE LOS PAGOS REALIZADOS POR LA  DETERMINACION DEL IMPUESTO SOBRE LA RENTA E IMPUESTO AL ACTIVO DIFERIDO POR DESCONSOLIDACION AL 31 DE DICIEMBRE DE 2013 Y EL PAGADO HASTA EL 30 DE ABRIL DEL 2018"},
            {"Anexo 20".ToUpper()               , "0"   ,"9"    ,""                     ,"20.- INVERSIONES"},
            {"Anexo 21".ToUpper()               , "0"   ,"12"   ,"Generales|C100"       ,"21.- CUENTAS Y DOCUMENTOS POR COBRAR Y POR PAGAR EN MONEDA EXTRANJERA"},
            {"Anexo 22".ToUpper()               , "0"   ,"25"   ,"Generales|C101"       ,"22.- PRESTAMOS DEL EXTRANJERO "},
            {"Anexo 23".ToUpper()               , "0"   ,"14"   ,"Generales|C61,C62"    ,"23.- INTEGRACION DE PERDIDAS FISCALES DE EJERCICIOS ANTERIORES"},
            {"CDF".ToUpper()                    , "78"  ,"5"    ,""                     ,"CUESTIONARIO DE DIAGNOSTICO FISCAL (REVISION DEL CONTADOR PUBLICO)"},
            {"MPT".ToUpper()                    , "111" ,"3"    ,""                     ,"CUESTIONARIO EN MATERIA DE PRECIOS DE TRANSFERENCIA (REVISION DEL CONTADOR PÚBLICO)"},
            {"Notas".ToUpper()                  , "48"  ,"1"    ,""                     ,""},
            {"Declaratoria".ToUpper()           , "45"  ,"1"    ,""                     ,""},
            {"Opinión".ToUpper()                , "45"  ,"1"    ,""                     ,""},
            {"Informe".ToUpper()                , "45"  ,"1"    ,""                     ,""},
            {"Información Adicional".ToUpper()  , "45"  ,"1"    ,""                     ,""}
        };

        public String[] _nombre;
        /// <summary>
        /// Arma el datagridView omitiendo los que no tienen descripcion
        /// </summary>
        /// <param name="_grilla"></param>
        public void _CargarGrilla(DataGridView _grilla)
        {
            //obtiene el numero de registros
            int numf = (_HojasSPR.Length) / _HojasSPR.GetLength(1);
            _grilla.Rows.Clear();
            for (int k = 0; k < numf; k++)
            {
                //Si la descripcion es vacio no lo agrega
                if (_HojasSPR[k, 4].ToString().Trim() != "")
                {
                    _grilla.Rows.Add(_HojasSPR[k, 0], _HojasSPR[k, 4], "false");
                }
            }
        }
        /// <summary>
        /// Prepara el excel para la impresion . oculta los indies sin datos de la grilla
        /// </summary>
        /// <param name="_Ocultar"></param> true muesta los datos false oculta
        /// <param name="_grilla"></param>hojas en las que se aplicara los cambios
        /// <param name="mostrar"></param>si se manda true pon visible todas las filas
        public void _PrepararImpresion(Boolean _Ocultar, DataGridView _grilla,Boolean mostrar)
        {
            if (!Verificar(_grilla))
            {
                return;
            }
            int EspacioFilas = 0;
            int fila = 3;
            int columna = 1;
            int ind;
            int fv = 0;
            //obtenemos el numero de hojas

            int numhojas = Globals.ThisAddIn.Application.Sheets.Count;
            String nom;
            //cargar array de nombres
            _Cargararraynombre(_HojasSPR);
            //Contraseña
            Generales.Proteccion(false);
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            int numhj = 0;
            for (int i = 1; i <= _grilla.RowCount; i++)
            {
                if (_grilla.Rows[i - 1].Cells["Imprimir"].Value.ToString().Trim().ToUpper() == "TRUE" || mostrar)
                {
                    numhj = Array.IndexOf(_nombre, _grilla.Rows[i - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper());
                    Globals.ThisAddIn.Application.Sheets[_grilla.Rows[i - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()].Activate();
                    fila = 3;
                    EspacioFilas = 0;
                    nom = Globals.ThisAddIn.Application.Sheets[_grilla.Rows[i - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()].Name.ToString().Trim();
                    ind = Array.IndexOf(_nombre, Globals.ThisAddIn.Application.Sheets[_grilla.Rows[i - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()].Name.ToString().Trim().ToUpper());
                    do
                    {
                        columna = 1;
                        fv = 0;
                        for (int j = 3; j <= _ValidarInt(_HojasSPR[numhj, 2]); j++)
                        {
                            if ((_ValidarString(((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[fila, j].Value).Trim() == "" || _ValidarInt(((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[fila, j].Value) == 0) && _ValidarString(((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[fila, 1].Value).Trim().Length > 0)
                            {
                                fv++;
                            }
                        }
                        if (!mostrar)
                        {
                            if (fv == _ValidarInt(_HojasSPR[numhj, 2]) - 2)
                            {
                                ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Rows[fila].Hidden = _Ocultar;
                            }
                            else
                            {
                                ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Rows[fila].Hidden = false;
                            }
                        }
                        else {
                            ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Rows[fila].Hidden = false;
                        }
                        

                        if (_ValidarString(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[fila, columna].Value).Trim().Length == 0 && _ValidarString(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[fila, columna + 1].Value).Trim().Length == 0)
                        {
                            EspacioFilas++;
                        }
                        else
                        {
                            EspacioFilas = 0;
                        }
                        fila++;

                        if (EspacioFilas == 12 && ind != -1)
                        {
                            _HojasSPR[ind, 1] = (fila - 12).ToString().Trim();
                        }
                    } while (EspacioFilas < 12);//Si existen mas de 12 espacios en blanco ya no genera mas filas
                }
            }
            Generales.Proteccion(true);
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }

        public String _ValidarString(object _val)
        {
            try { return Convert.ToString(_val); }
            catch { return ""; }
        }
        public void _Cargararraynombre(string[,] _val)
        {
            int numf = (_val.Length) / _val.GetLength(1);
            _nombre = new String[numf];
            for (int k = 0; k < numf; k++)
            {
                _nombre[k] = _val[k, 0];
            }
        }
        public int _ValidarInt(object _val)
        {
            try { return Convert.ToInt32(_val); }
            catch { return 0; }
        }

        /// <summary>
        /// Prepara la vista previa
        /// </summary> 
        /// <param name="_grilla" ></param>El DataGridView de donde se verifica lo que se imprime
        /// <param name="_BandW"></param>Ture si es en Blanco y Negro y False si es a color
        /// <param name="_impresora"></param>nombre de la impresora o Type.Missing si mostraremos una vista previa
        public void _Imprimir(DataGridView _grilla, Boolean _BandW, object _impresora)
        {
            if (!Verificar(_grilla))
            {
                MessageBox.Show("Debe seleccionar al menos un anexo. ", "Imprimir SIPRED", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Excel.Workbook libron = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Workbook libro = libron;

            Generales.Proteccion(false);
            //Desactivamos los mensajes de Alerta del Excel
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            int numhj = 0;
            _Cargararraynombre(_HojasSPR);

            for (int k = 1; k <= _grilla.RowCount; k++)
            {
                numhj = Array.IndexOf(_nombre, _grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper());
                if (_grilla.Rows[k - 1].Cells["Imprimir"].Value.ToString().Trim().ToUpper() == "TRUE")
                {
                    numhj = Array.IndexOf(_nombre, _grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper());
                    //para mantener las dos primeras filas y columnas fijas en la vista previa
                    ((Excel.Worksheet)libro.Sheets[_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()]).PageSetup.PrintTitleRows = "$1:$2";
                    ((Excel.Worksheet)libro.Sheets[_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()]).PageSetup.PrintTitleColumns = "$A:$B";
                    ((Excel.Worksheet)libro.Sheets[_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()]).PageSetup.Zoom = 65;
                    ((Excel.Worksheet)libro.Sheets[_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()]).PageSetup.BlackAndWhite = _BandW;
                    //Cuando Es un Anexo: Orientacion horizontal
                    if (_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper().Contains("ANEXO"))
                    {
                        ((Excel.Worksheet)libro.Sheets[_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()]).PageSetup.Orientation = XlPageOrientation.xlLandscape;
                    }
                    else
                    {
                        ((Excel.Worksheet)libro.Sheets[_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()]).PageSetup.Orientation = XlPageOrientation.xlPortrait;
                    }
                }
                else if (numhj != -1)
                {
                    ((Excel.Worksheet)libro.Sheets[_grilla.Rows[k - 1].Cells["Anexo"].Value.ToString().Trim().ToUpper()]).Visible = XlSheetVisibility.xlSheetHidden;
                }

                //Notas a ocultar
                //Si se ocultaron hojas las vuelve visible todas
                for (int x = 1; x <= libro.Worksheets.Count; x++)
                {
                    if (Array.IndexOf(_nombre, ((Excel.Worksheet)libro.Sheets[x]).Name.ToString().Trim().ToUpper()) != -1 && _HojasSPR[Array.IndexOf(_nombre, ((Excel.Worksheet)libro.Sheets[x]).Name.ToString().Trim().ToUpper()), 4].Trim().Length == 0)
                    {
                        ((Excel.Worksheet)libro.Sheets[x]).Visible = XlSheetVisibility.xlSheetHidden;
                    }

                }

            }
            //Generales.Proteccion(false);
            if (_impresora == Type.Missing)
            {
                libro.PrintOut(Type.Missing, Type.Missing, Type.Missing, true, _impresora, Type.Missing, Type.Missing, Type.Missing);
            }
            else if (_impresora.ToString() == "PDF")
            {
                string _Path = Configuration.Path + "\\SIPRED" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".pdf";
                libro.ExportAsFixedFormat(Type: XlFixedFormatType.xlTypePDF, Filename: _Path, Quality: XlFixedFormatQuality.xlQualityStandard, OpenAfterPublish: true);
            }
            else
            {
                libro.PrintOut(Type.Missing, Type.Missing, Type.Missing, false, _impresora, Type.Missing, Type.Missing, Type.Missing);
            }

            //Si se ocultaron hojas las vuelve visible todas
            for (int k = 1; k <= libro.Worksheets.Count; k++)
            {
                if (Array.IndexOf(_nombre, ((Excel.Worksheet)libro.Sheets[k]).Name.ToString().Trim().ToUpper()) != -1)
                {
                    ((Excel.Worksheet)libro.Sheets[k]).Visible = XlSheetVisibility.xlSheetVisible;
                }
                else
                {
                    ((Excel.Worksheet)libro.Sheets[k]).Visible = XlSheetVisibility.xlSheetHidden;
                }

            }
            Generales.Proteccion(true);
            //Activamos los mensajes de Alerta del Excel
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }

        public Boolean Verificar(DataGridView _grilla)
        {
            Boolean resp = false;
            for (int k = 1; k <= _grilla.RowCount; k++)
            {
                if (_grilla.Rows[k - 1].Cells["Imprimir"].Value.ToString().Trim().ToUpper() == "TRUE")
                {
                    resp = true;

                }
            }
            return resp;
        }

        public void Cerrar() {

            Excel.Workbook libro = Globals.ThisAddIn.Application.ActiveWorkbook;

            Generales.Proteccion(false);
            //Activamos los mensajes de Alerta del Excel
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            //Si se ocultaron hojas las vuelve visible todas
            for (int k = 1; k <= libro.Worksheets.Count; k++)
            {
                if (Array.IndexOf(_nombre, ((Excel.Worksheet)libro.Sheets[k]).Name.ToString().Trim().ToUpper()) != -1)
                {
                    ((Excel.Worksheet)libro.Sheets[k]).Visible = XlSheetVisibility.xlSheetVisible;
                }
                else
                {
                    ((Excel.Worksheet)libro.Sheets[k]).Visible = XlSheetVisibility.xlSheetHidden;
                }

            }
            Generales.Proteccion(true);
            //Activamos los mensajes de Alerta del Excel
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }


    }
}
