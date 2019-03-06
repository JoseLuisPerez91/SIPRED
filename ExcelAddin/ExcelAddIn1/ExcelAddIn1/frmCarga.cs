using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using ExcelAddIn.Access;
using System.IO;

namespace ExcelAddIn1
{
    public partial class frmCarga : Form
    {
        public bool _ProcessJson;
        public frmCarga()
        {
            InitializeComponent();
        }
        private void frmCarga_Load(object sender, EventArgs e)
        {
            string _Path = Configuration.Path;
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            this._ProcessJson = false;
            this.Visible = true;

            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\TiposPlantillas.json"))
                {
                    if (_Connection)
                    {
                        KeyValuePair<bool, System.Data.DataTable> _TipoPlantilla = new lSerializados().ObtenerUpdate();

                        foreach (DataRow _Row in _TipoPlantilla.Value.Rows)
                        {
                            string _IdTipoPlantilla = _Row["IdTipoPlantilla"].ToString();
                            string _Fecha_Modificacion = _Row["Fecha_Modificacion"].ToString();
                            string _Linea = null;

                            if (File.Exists(_Path + "\\jsons\\Update" + _IdTipoPlantilla + ".txt"))
                            {
                                StreamReader sw = new StreamReader(_Path + "\\Jsons\\Update" + _IdTipoPlantilla + ".txt");
                                _Linea = sw.ReadLine();
                                sw.Close();

                                if (_Linea != null)
                                {
                                    if (_Linea != _Fecha_Modificacion)
                                    {
                                        this.TopMost = false;
                                        this.Enabled = false;
                                        this.Hide();
                                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                                        _FileJsonfrm._Form = this;
                                        _FileJsonfrm._Process = false;
                                        _FileJsonfrm._Update = true;
                                        _FileJsonfrm._window = this.Text;
                                        _FileJsonfrm.Show();
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (!_Connection)
                    {
                        MessageBox.Show(_Message.Replace("...", ", para crear los archivos base..."), "Creación de Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.btnAccept.Enabled = false;
                        return;
                    }
                    else
                    {
                        this._ProcessJson = true;
                        this.TopMost = false;
                        this.Enabled = false;
                        this.Hide();
                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                        _FileJsonfrm._Form = this;
                        _FileJsonfrm._Process = false;
                        _FileJsonfrm._Update = false;
                        _FileJsonfrm._window = this.Text;
                        _FileJsonfrm.Show();
                        return;
                    }
                }
            }
            else
            {
                if (!Directory.Exists(_Path + "\\jsons"))
                {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if (!Directory.Exists(_Path + "\\templates"))
                {
                    Directory.CreateDirectory(_Path + "\\templates");
                }

                this._ProcessJson = true;
                this.TopMost = false;
                this.Enabled = false;
                this.Hide();
                FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                _FileJsonfrm._Form = this;
                _FileJsonfrm._Process = false;
                _FileJsonfrm._Update = false;
                _FileJsonfrm._window = this.Text;
                _FileJsonfrm.Show();
                return;
            }

            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);

            if (_Excel.Extension != ".xlsm")
            {
                MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }

            QuitarFormulas();
        }
        private void frmCarga_Shown(object sender, EventArgs e)
        {
            //_f_run_background();
        }
        private void btnAccept_Click(object sender, EventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
        }
        #region Variables
        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

        string[,] HojasSPR = new string[,] {
            {"Contribuyente".ToUpper()          , "31"  ,"3"    ,""                     },
            {"Contador".ToUpper()               , "35"  ,"3"    ,""                     },
            {"Representante".ToUpper()          , "36"  ,"3"    ,""                     },
            {"Generales".ToUpper()              , "446" ,"3"    ,""                     },
            {"Anexo 1".ToUpper()                , "0"   ,"10"   ,""                     },
            {"Anexo 2".ToUpper()                , "0"   ,"9"    ,""                     },
            {"Anexo 3".ToUpper()                , "0"   ,"22"   ,""                     },
            {"Anexo 4".ToUpper()                , "0"   ,"5"    ,""                     },
            {"Anexo 5".ToUpper()                , "0"   ,"14"   ,""                     },
            {"Anexo 6".ToUpper()                , "0"   ,"5"    ,""                     },
            {"Anexo 7".ToUpper()                , "0"   ,"37"   ,""                     },
            {"Anexo 8".ToUpper()                , "0"   ,"9"    ,""                     },
            {"Anexo 9".ToUpper()                , "0"   ,"9"    ,""                     },
            {"Anexo 10".ToUpper()               , "0"   ,"15"   ,""                     },
            {"Anexo 11".ToUpper()               , "0"   ,"4"    ,""                     },
            {"Anexo 12".ToUpper()               , "0"   ,"13"   ,"Generales|C96"        },
            {"Anexo 13".ToUpper()               , "0"   ,"10"   ,"Generales|C97"        },
            {"Anexo 14".ToUpper()               , "0"   ,"12"   ,""                     },
            {"Anexo 15".ToUpper()               , "0"   ,"4"    ,""                     },
            {"Anexo 16".ToUpper()               , "0"   ,"11"   ,"Generales|C57"        },
            {"Anexo 17".ToUpper()               , "0"   ,"4"    ,"Generales|C57"        },
            {"Anexo 18".ToUpper()               , "0"   ,"4"    ,""                     },
            {"Anexo 19".ToUpper()               , "0"   ,"7"    ,"Generales|C98"        },
            {"Anexo 20".ToUpper()               , "0"   ,"9"    ,""                     },
            {"Anexo 21".ToUpper()               , "0"   ,"12"   ,"Generales|C100"       },
            {"Anexo 22".ToUpper()               , "0"   ,"25"   ,"Generales|C101"       },
            {"Anexo 23".ToUpper()               , "0"   ,"14"   ,"Generales|C61,C62"    },
            {"CDF".ToUpper()                    , "78"  ,"5"    ,""                     },
            {"MPT".ToUpper()                    , "111" ,"3"    ,""                     },
            {"Notas".ToUpper()                  , "48"  ,"1"    ,""                     },
            {"Declaratoria".ToUpper()           , "45"  ,"1"    ,""                     },
            {"Opinión".ToUpper()                , "45"  ,"1"    ,""                     },
            {"Informe".ToUpper()                , "45"  ,"1"    ,""                     },
            {"Información Adicional".ToUpper()  , "45"  ,"1"    ,""                     }
        };

        String[] nombre;
        #endregion
        public void QuitarFormulas()
        {
            //objeto vacio
            object obj = Type.Missing;
            int numhojas = 0;
            int EspacioFilas = 0;
            int EspacioColumnas = 0;
            int fila = 1;
            int columna = 1;
            int ind = 0;
            String nom = "";
            String pr = "";
            Microsoft.Office.Interop.Excel.Range cell1;
            Microsoft.Office.Interop.Excel.Range cell2;
            Microsoft.Office.Interop.Excel.Range range;
            double num = 0;
            String psw = "";

            //Nuevo Excel
            Excel.Application exceln = new Excel.Application();
            //libro abierto
            Excel.Workbook libro = Globals.ThisAddIn.Application.ActiveWorkbook;
            //nuevo libro
            //Excel.Workbook libron = exceln.Workbooks.Add(obj);
            Excel.Workbook libron = libro;
            //obtenemos el numero de hojas
            numhojas = libro.Sheets.Count;
            //cargar array de nombres
            Cargararraynombre(HojasSPR);
            //seleccionamos instanci hoja
            Excel.Worksheet hoja = libro.Sheets[1];
            //Creamos instancia nueva hoja
            Excel.Worksheet hojan = new Excel.Worksheet();
            int[] hojas = new int[numhojas];
            //Contraseña
            //psw = "AAAABABABAAG";
            //LENARHOJAS
            Globals.ThisAddIn.Application.DisplayAlerts = false;

            Generales.Proteccion(false);

            for (int i = 1; i <= numhojas; i++)
            {
                //seleccionamos la hoja de orden i
                hoja = libro.Sheets[i];
                //Creamos nueva hoja de orden i
                hojan = libron.Sheets[i];
                hojas[i - 1] = ValidarInt(Regex.Replace(hoja.Name, @"[^\d]", "").ToString().Trim());
                //seleccionamos la hoja numero i
                hoja.Activate();
                hojan.Activate();

                nom = libron.Worksheets[i].Name.ToString().Trim();
                ind = Array.IndexOf(nombre, libron.Worksheets[i].Name.ToString().Trim().ToUpper());
                //Barra de progreso
                if (this == null) return;
                Invoke(new System.Action(() => this.label1.Text = "Trabajando Hoja : [" + (Globals.ThisAddIn.Application.ActiveSheet).Name + "] .........."));
                if (this == null) return;
                Invoke(new System.Action(() => pgb_proceso.Value = pgb_proceso.Value + pgb_proceso.Step));

                //pasamos los datos
                EspacioFilas = 0;
                fila = 1;
                columna = 1;
                //hojan.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);

                do
                {
                    //Eliminar Comentarios
                    foreach (Comment a in ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Comments)
                    {
                        a.Delete();
                    }

                    columna = 1;
                    EspacioColumnas = 0;
                    do
                    {
                        cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[fila, columna];
                        if (ValidarString(cell1.FormulaLocal).Trim().Length > 0)
                        {
                            if (ValidarString(cell1.FormulaLocal).Trim().Substring(0, 1) == "=")
                            {
                                hojan.Cells[fila, columna].Value = ValidarString(hoja.Cells[fila, columna].Value);
                            }
                        }

                        if (Double.TryParse(ValidarString(cell1.Value), out num))
                        {
                            if (ValidarString(hoja.Cells[fila, columna]).Trim().Substring(0, 1) != "0")
                            {
                                try
                                {
                                    //Evitar el error de seguridad
                                    if (!hojan.Cells[fila, columna].Locked)
                                    {
                                        hojan.Cells[fila, columna].Value = num;
                                    }

                                }
                                catch { }
                            }
                        }

                        if (ind != -1)
                        {
                            if (ValidarInt(ValidarString(HojasSPR[ind, 2]).Trim()) <= columna && ((ValidarInt(ValidarString(HojasSPR[ind, 1]).Trim()) <= fila) || (ValidarInt(ValidarString(HojasSPR[ind, 1]).Trim()) == 0)))
                            {
                                hojan.Cells[fila, columna].Value = "";
                            }
                        }

                        if (ValidarString(hoja.Cells[fila, columna].Value).Trim().Length == 0)
                        {
                            EspacioColumnas++;
                        }
                        else
                        {
                            EspacioColumnas = 0;
                        }
                        columna++;

                    } while (EspacioColumnas < 10);//Si existen mas de 10 espacios en blanco ya no genera mas columnas

                    columna = 1;
                    pr = ValidarString(hoja.Cells[fila, columna].Value).Trim();
                    if (ValidarString(hoja.Cells[fila, columna].Value).Trim().Length == 0 && ValidarString(hoja.Cells[fila, columna + 1].Value).Trim().Length == 0)
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
                        HojasSPR[ind, 1] = (fila - 12).ToString().Trim();
                    }
                } while (EspacioFilas < 12);//Si existen mas de 12 espacios en blanco ya no genera mas filas
            }
            //Ordenar Hojas Excel

            //deactivar mensajes alerta que genera al eliminar
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            for (int j = 0; j < 5; j++)
            {
                for (int i = 1; i <= numhojas; i++)
                {
                    try
                    {
                        nom = libron.Worksheets[i].Name.ToString().Trim();
                        ind = Array.IndexOf(nombre, libron.Worksheets[i].Name.ToString().Trim().ToUpper());
                        //Barra de progreso
                        if (this == null) return;
                        Invoke(new System.Action(() => this.label1.Text = "Trabajando Hoja : [" + (Globals.ThisAddIn.Application.ActiveSheet).Name + "] .........."));

                        if (ind != -1)
                        {
                            ind++;
                            libron.Sheets[i].Activate();
                            //Color plomo de la etiqueta
                                //Protegerhoja
                                //((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Tab.Color = System.Drawing.Color.FromArgb(251, 155, 13);
                                //((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Protect(psw, true);
                            
                                ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Tab.Color = System.Drawing.Color.FromArgb(100, 100, 100);
                           
                            //((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Tab.ColorIndex = 0;
                            ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Move(Globals.ThisAddIn.Application.Worksheets[ind]);

                            //Ocultar columnas

                            cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[80, ValidarInt(ValidarString(HojasSPR[i - 1, 2]).Trim())];
                            cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                            range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                            range.EntireColumn.Hidden = true;

                            cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                            cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[80, ValidarInt(ValidarString(HojasSPR[i - 1, 2]).Trim())];
                            range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                            ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Columns.Hidden = false;

                            //Ocultar filas
                            cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[ValidarInt(ValidarString(HojasSPR[i - 1, 1]).Trim()), 1];
                            cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                            range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                            range.EntireRow.Hidden = true;

                            cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                            cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[ValidarInt(ValidarString(HojasSPR[i - 1, 1]).Trim()), ValidarInt(ValidarString(HojasSPR[i - 1, 2]).Trim())];
                            range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                            ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Rows.Hidden = false;
                        }
                        else
                        {

                            Excel.Worksheet m_objSheet = (Excel.Worksheet)(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.get_Item(i));
                            m_objSheet.Delete();
                            numhojas--;
                            i = i - 1;
                        }
                    }
                    catch (Exception e)
                    {
                        Generales.Proteccion(true);
                        MessageBox.Show(e.Message);
                    }
                }
            }

            //Agregar Hojas
            for (int i = 0; i < 5; i++)
            {
                int count = libron.Worksheets.Count;
                Excel.Worksheet HN = libron.Worksheets.Add(Type.Missing,
                        libron.Worksheets[count], Type.Missing, Type.Missing);
                if (i == 0)
                {
                    HN.Name = "Notas";
                    HN.Cells[1, 1].Value = "SERVICIO DE ADMINISTRACION TRIBUTARIA";
                    HN.Cells[1, 1].Font.Size = 12;
                    HN.Cells[3, 1].Value = "SISTEMA DE PRESENTACION DEL DICTAMEN 2017";
                    HN.Cells[3, 1].Font.Size = 10;
                    HN.Cells[5, 1].Value = "NOMBRE DEL CONTRIBUYENTE:";
                    HN.Cells[5, 1].Font.Size = 9;
                    HN.Cells[5, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    HN.Cells[6, 1].AddComment("Es réplica de: \n Anexo: Contribuyente \n Índice: 01A001000 \n Columna: C");
                    HN.Cells[6, 1].Value = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets["Contribuyente"]).Range["C4"].Value;
                    HN.Cells[8, 1].Value = "INFORMACION DEL ANEXO : 4.1. NOTAS A LOS ESTADOS FINANCIEROS";
                    HN.Cells[8, 1].Font.Size = 9;
                    HN.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    HN.Cells[46, 1].Value = "LISTA DE NOTAS:";
                    HN.Cells[46, 1].Font.Size = 9;
                    HN.Cells[46, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    range = HN.Range[HN.Cells[12, 1], HN.Cells[44, 1]];


                    range.Merge();

                    range.EntireColumn.ColumnWidth = 100;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;


                    range = HN.Range[HN.Cells[1, 1], HN.Cells[1, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[44, 1]];
                    range.Font.Name = "Arial";
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[3, 1]];
                    range.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.Font.Bold = true;


                    //Ocultar columnas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[48, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireColumn.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[47, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Columns.Hidden = false;

                    //Ocultar filas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[48, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireRow.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[47, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Rows.Hidden = false;

                }
                if (i == 1)
                {
                    HN.Name = "Declaratoria";

                    HN.Cells[1, 1].Value = "SERVICIO DE ADMINISTRACION TRIBUTARIA";
                    HN.Cells[1, 1].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[1, 1].Font.Size = 12;
                    HN.Cells[3, 1].Value = "SISTEMA DE PRESENTACION DEL DICTAMEN 2017";
                    HN.Cells[3, 1].Font.Size = 10;
                    HN.Cells[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[5, 1].Value = "NOMBRE DEL CONTRIBUYENTE:";
                    HN.Cells[5, 1].Font.Size = 9;
                    HN.Cells[5, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    HN.Cells[6, 1].AddComment("Es réplica de: \n Anexo: Contribuyente \n Índice: 01A001000 \n Columna: C");
                    HN.Cells[6, 1].Value = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets["Contribuyente"]).Range["C4"].Value;
                    HN.Cells[8, 1].Value = "INFORMACION DEL ANEXO : 9.1. DECLARATORIA";
                    HN.Cells[8, 1].Font.Size = 9;
                    HN.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    range = HN.Range[HN.Cells[13, 1], HN.Cells[45, 1]];

                    range.Merge();

                    range.EntireColumn.ColumnWidth = 100;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;



                    range = HN.Range[HN.Cells[1, 1], HN.Cells[1, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[44, 1]];
                    range.Font.Name = "Arial";
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[3, 1]];
                    range.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.Font.Bold = true;

                    //Ocultar columnas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireColumn.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Columns.Hidden = false;

                    //Ocultar filas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireRow.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Rows.Hidden = false;

                }
                if (i == 2)
                {
                    HN.Name = "Opinión";

                    HN.Cells[1, 1].Value = "SERVICIO DE ADMINISTRACION TRIBUTARIA";
                    HN.Cells[1, 1].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[1, 1].Font.Size = 12;
                    HN.Cells[3, 1].Value = "SISTEMA DE PRESENTACION DEL DICTAMEN 2017";
                    HN.Cells[3, 1].Font.Size = 10;
                    HN.Cells[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[5, 1].Value = "NOMBRE DEL CONTRIBUYENTE:";
                    HN.Cells[5, 1].Font.Size = 9;
                    HN.Cells[5, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    HN.Cells[6, 1].AddComment("Es réplica de: \n Anexo: Contribuyente \n Índice: 01A001000 \n Columna: C");
                    HN.Cells[6, 1].Value = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets["Contribuyente"]).Range["C4"].Value;
                    HN.Cells[8, 1].Value = "INFORMACION DEL ANEXO : OPINION";
                    HN.Cells[8, 1].Font.Size = 9;
                    HN.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    range = HN.Range[HN.Cells[13, 1], HN.Cells[45, 1]];

                    range.Merge();

                    range.EntireColumn.ColumnWidth = 100;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[1, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[44, 1]];
                    range.Font.Name = "Arial";
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[3, 1]];
                    range.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.Font.Bold = true;

                    //Ocultar columnas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireColumn.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Columns.Hidden = false;

                    //Ocultar filas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireRow.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Rows.Hidden = false;

                }
                if (i == 3)
                {
                    HN.Name = "Informe";

                    HN.Cells[1, 1].Value = "SERVICIO DE ADMINISTRACION TRIBUTARIA";
                    HN.Cells[1, 1].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[1, 1].Font.Size = 12;
                    HN.Cells[3, 1].Value = "SISTEMA DE PRESENTACION DEL DICTAMEN 2017";
                    HN.Cells[3, 1].Font.Size = 10;
                    HN.Cells[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[5, 1].Value = "NOMBRE DEL CONTRIBUYENTE:";
                    HN.Cells[5, 1].Font.Size = 9;
                    HN.Cells[5, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    HN.Cells[6, 1].AddComment("Es réplica de: \n Anexo: Contribuyente \n Índice: 01A001000 \n Columna: C");
                    HN.Cells[6, 1].Value = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets["Contribuyente"]).Range["C4"].Value;
                    HN.Cells[8, 1].Value = "INFORMACION DEL ANEXO : INFORME";
                    HN.Cells[8, 1].Font.Size = 9;
                    HN.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    range = HN.Range[HN.Cells[13, 1], HN.Cells[45, 1]];

                    range.Merge();

                    range.EntireColumn.ColumnWidth = 100;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;



                    range = HN.Range[HN.Cells[1, 1], HN.Cells[1, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[44, 1]];
                    range.Font.Name = "Arial";
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[3, 1]];
                    range.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.Font.Bold = true;

                    //Ocultar columnas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireColumn.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Columns.Hidden = false;

                    //Ocultar filas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireRow.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Rows.Hidden = false;

                }
                if (i == 4)
                {
                    HN.Name = "Información Adicional";

                    HN.Cells[1, 1].Value = "SERVICIO DE ADMINISTRACION TRIBUTARIA";
                    HN.Cells[1, 1].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[1, 1].Font.Size = 12;
                    HN.Cells[3, 1].Value = "SISTEMA DE PRESENTACION DEL DICTAMEN 2017";
                    HN.Cells[3, 1].Font.Size = 10;
                    HN.Cells[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    HN.Cells[5, 1].Value = "NOMBRE DEL CONTRIBUYENTE:";
                    HN.Cells[5, 1].Font.Size = 9;
                    HN.Cells[5, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    HN.Cells[6, 1].AddComment("Es réplica de: \n Anexo: Contribuyente \n Índice: 01A001000 \n Columna: C");
                    HN.Cells[6, 1].Value = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets["Contribuyente"]).Range["C4"].Value;
                    HN.Cells[8, 1].Value = "INFORMACION DEL ANEXO : INFORMACION ADICIONAL";
                    HN.Cells[8, 1].Font.Size = 9;
                    HN.Cells[8, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    range = HN.Range[HN.Cells[13, 1], HN.Cells[45, 1]];

                    range.Merge();

                    range.EntireColumn.ColumnWidth = 100;
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[1, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[44, 1]];
                    range.Font.Name = "Arial";
                    range = HN.Range[HN.Cells[1, 1], HN.Cells[3, 1]];
                    range.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range = HN.Range[HN.Cells[3, 1], HN.Cells[3, 1]];
                    range.Font.Bold = true;

                    //Ocultar columnas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireColumn.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Columns.Hidden = false;

                    //Ocultar filas
                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[46, 2];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1048576, 16384];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    range.EntireRow.Hidden = true;

                    cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                    cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[45, 1];
                    range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Range[cell1, cell2].Rows.Hidden = false;

                }
            }

            //for (int i = 1; i <= numhojas; i++)
            //{
            //    try
            //    {
            //        nom = libron.Worksheets[i].Name.ToString().Trim();
            //        ind = Array.IndexOf(nombre, libron.Worksheets[i].Name.ToString().Trim().ToUpper());
            //        if (ind != -1)
            //        {
            //            ind++;
            //            libron.Sheets[i].Activate();
            //            //((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Protect(ExcelAddIn.Access.Configuration.PwsExcel, true);

            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        Generales.Proteccion(true);
            //        MessageBox.Show(e.Message);
            //    }
            //}
            Generales.Proteccion(true);


            if (this == null) return;
            Invoke(new System.Action(() => this.pgb_proceso.Value = this.pgb_proceso.Maximum));
            Invoke(new System.Action(() => this.label1.Text = "Proceso Terminado."));

            GuardarExcel();

            //activar mensajes alerta
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            Invoke(new System.Action(() => this.Visible = false));
            //this.Visible = false;
            Invoke(new System.Action(() => this.Close()));
        }
        public void Cargararraynombre(string[,] val)
        {
            int numf = (val.Length) / val.GetLength(1);
            nombre = new String[numf];
            for (int k = 0; k < numf; k++)
            {
                nombre[k] = val[k, 0];
            }
        }
        public void GuardarExcel()
        {
            //guardar nuevo libro
            object obj = Type.Missing;
            Excel.Workbook libron = Globals.ThisAddIn.Application.ActiveWorkbook;
            SaveFileDialog1 = new SaveFileDialog()
            {
                DefaultExt = "*.xlsx",
                //SaveFileDialog1.FileName = Globals.ThisAddIn.Application.ActiveWorkbook.Name + ".xls";
                FileName = libron.Name + ".xlsx",
                Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
            };
            if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                libron.SaveAs(SaveFileDialog1.FileName, Excel.XlFileFormat.xlOpenXMLWorkbook, obj, obj, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, obj, obj, obj);
            }
        }
        public String ValidarString(object val)
        {
            try { return Convert.ToString(val); }
            catch { return ""; }
        }
        public int ValidarInt(object val)
        {
            try { return Convert.ToInt32(val); }
            catch { return 0; }
        }
        public void MensageBloqueo(Excel.Worksheet Sh)
        {
            String CondCad = "";
            string[] arg;
            string[] cond;
            Boolean res = true;
            String Vcon = "";
            //cargar array de nombres
            Cargararraynombre(HojasSPR);

            String nom = Sh.Name.ToString().Trim();
            int ind = Array.IndexOf(nombre, Sh.Name.ToString().Trim().ToUpper());
            if (ind != -1)
            {
                //ind++;
                Sh.Activate();
                //Color plomo de la etiqueta
                //if (ind == 15 || ind == 16 || ind == 19 || ind == 20 || ind == 21 || ind == 23 || ind == 24 || ind == 25)
                //{
                if (HojasSPR[ind, 3].Trim().Length > 0)
                {
                    //Capturo la condicion
                    CondCad = HojasSPR[ind, 3].Trim();
                    arg = CondCad.Split('|');
                    nom = arg[0].ToString().Trim();
                    ind = Array.IndexOf(nombre, nom.ToUpper());
                    cond = arg[1].ToString().Trim().Split(',');

                    foreach (string i in cond)
                    {
                        Vcon = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets["Generales"]).Range[i].Formula;

                        if (Vcon == "SI" || Vcon == "si")
                        {
                            res = false;
                            break;
                        }
                        if (Vcon == "NO" || Vcon == "no")
                        {
                            res = true;
                        }
                    }
                    if (res)
                    {
                        MessageBox.Show("No es posible seleccionar el anexo debido a que se encuentra deshabilitado.", "SPRIND", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[Sh.Index - 1]).Activate();
                    }
                }
            }
        }
        private void frmCarga_FormClosed(object sender, FormClosedEventArgs e)
        {
            //if (this._ProcessJson == false)
            //{
            //    GuardarExcel();
            //    //Ribbon2 ai = new Ribbon2();

            //    //ai._Form = this;
            //    //ai.GuardarExcel();
            //}
        }
    }
}
