﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class Ribbon2
    {

        int NroFilaPrincipal = 0;
        int NroColPrincipal = 0;
        bool tieneformula = false;



        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnNew_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnNewIndice_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet ActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range objRange;

            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            NroFilaPrincipal = currentCell.Row;
            NroColPrincipal = currentCell.Column;
            int iTotalColumns; int k = 1;
            bool puedeinsertar = false;
            string IndicePrevio;

            string tag = "";
            try
            {

                objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, 1];
                IndicePrevio = objRange.get_Value(Type.Missing).ToString();
                if (IndicePrevio.ToUpper().Trim() != "EXPLICACION")
                {
                    foreach (Excel.Name item in wb.Names)
                    {
                        tag = item.RefersToRange.Cells.get_Address();
                        //Compare the active cell address with named range address
                        if (tag == objRange.Address)
                        {

                            if ((NroFilaPrincipal - 1) > 0)
                            {
                                objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal - 1, 1]; //Verifico que la fila anterior tenga formula
                                iTotalColumns = ActiveWorksheet.UsedRange.Columns.Count;

                                while (k <= iTotalColumns)
                                {

                                    if (objRange.Cells[k].HasFormula)
                                        tieneformula = true;

                                    k = k + 1;
                                }
                            }
                            puedeinsertar = true;


                            break;
                        }
                    }
                    if (puedeinsertar)
                    {
                        frmNewIndices NewIndices = new frmNewIndices(NroFilaPrincipal, tieneformula);
                        NewIndices.ShowDialog();

                    }
                    else
                    {
                        objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, 2];
                        var ConceptoPrevio = objRange.get_Value(Type.Missing);
                        if (ConceptoPrevio != null)
                        {
                            ConceptoPrevio = ConceptoPrevio.ToString();
                            if (ConceptoPrevio.Length >= 4)
                            {
                                if (ConceptoPrevio.Substring(0, 4).ToUpper() == "OTRO")
                                {
                                    NroFilaPrincipal = objRange.Row;
                                    NroColPrincipal = objRange.Column;
                                    if ((NroFilaPrincipal - 1) > 0)
                                    {
                                        // objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal - 1, Type.Missing]; //Verifico que la fila anterior tenga formula
                                        var RangeConFr = ActiveWorksheet.get_Range(string.Format("{0}:{0}", NroFilaPrincipal - 1, Type.Missing));
                                        iTotalColumns = ActiveWorksheet.UsedRange.Columns.Count;

                                        while (k <= iTotalColumns)
                                        {

                                            if (RangeConFr.Cells[k].HasFormula)
                                                tieneformula = true;

                                            k = k + 1;
                                        }
                                    }

                                    frmNewIndices NewIndices = new frmNewIndices(NroFilaPrincipal, tieneformula);
                                    NewIndices.ShowDialog();
                                }
                                else
                                    MessageBox.Show("No es posible agregar índices debajo del índice " + IndicePrevio, "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }


                    }
                }
                else {
                    MessageBox.Show("No es posible agregar índices debajo del índice EXPLICACION", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show("No es posible agregar índices en la fila seleccionada", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //MessageBox.Show(ex.Message);

            }

        }

        public static void AddNamedRange(int row, int col, string myrango)
        {
            Microsoft.Office.Tools.Excel.NamedRange NamedRange1;

            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);


            Excel.Range cell = worksheet.Cells[row, col];

            NamedRange1 = worksheet.Controls.AddNamedRange(cell, myrango);



        }

        private void btnDelIndice_Click(object sender, RibbonControlEventArgs e)
        {


            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            string IndiceActivo = "";
            string IndiceSiguiente = "";
            bool Eliminar = false;
            List<string> NombreRangos = new List<string>();
            List<string> NombreRangosDEL = new List<string>();
            List<int> FilaPadre = new List<int>();
            int FilapadreAux = 0;
            long dif = 0;
            string NamedRange = "";
            bool tienedif = false;
            Excel.Range objRange = null;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.Selection; // filas seleccionadas

            
            try
            {

                foreach (Excel.Range cell in currentCell.Cells)
                {

                    try
                    {


                        foreach (Excel.Name item1 in wb.Names)
                        {
                            // comparo la direccion de la celda con la del nombre del rango
                            if (item1.RefersToRange.Cells.get_Address() == cell.Address)
                            {
                                NamedRange = item1.Name;

                                break;
                            }
                        }

                        FilapadreAux = cell.Row;

                        if (!FilaPadre.Contains(FilapadreAux))
                            FilaPadre.Add(FilapadreAux);


                        objRange = (Excel.Range)sheet.Cells[cell.Row, 1];
                        IndiceActivo = objRange.Value2;

                      
                        if (IndiceActivo.ToUpper().Trim()=="EXPLICACION")
                        {
                            MessageBox.Show("No es posible eliminar el índice EXPLICACION." , "Eliminar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Eliminar = false;
                            break;
                        }
                        if ((NamedRange != "IA_" + IndiceActivo) || (NamedRange == ""))
                        {
                            MessageBox.Show("No es posible eliminar un índice de formato guía", "Eliminar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Eliminar = false;
                            break;
                        }
                        else
                        {
                            Eliminar = true;

                            NombreRangosDEL.Add("IA_" + IndiceActivo);
                        }

                        
                    }
                    catch (Exception ex)
                    {

                        //MessageBox.Show(ex.Message);

                    }

                }

                if (Eliminar)
                {


                    currentCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                   

                    objRange = (Excel.Range)sheet.Cells[currentCell.Cells.Row + 1, 1];
                    IndiceSiguiente = objRange.Value2;
                    if (IndiceSiguiente.ToUpper().Trim()=="EXPLICACION")
                        objRange.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    currentCell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    NombreRangosDEL.Sort();
                    string NM = NombreRangosDEL.FirstOrDefault();
                    sheet.Controls.Remove(NM);
                    foreach (Excel.Name item2 in wb.Names)
                        NombreRangos.Add(item2.Name);
                    string[] split = NM.Split('_');
                    NM = split[1];
                    // foreach (string Nm in NombreRangosDEL)
                    long NamedRng = Convert.ToInt64(NM) + 100;
                    string IndiceSig = "0" + Convert.ToString(NamedRng);
                    while (NombreRangos.Contains("IA_" + IndiceSig))
                    {
                        sheet.Controls.Remove("IA_" + IndiceSig);

                        NamedRng = Convert.ToInt64(IndiceSig) + 100;
                        IndiceSig = "0" + Convert.ToString(NamedRng);
                    }
                    FilaPadre.Sort();
                    int row = FilaPadre.FirstOrDefault();

                    objRange = (Excel.Range)sheet.Cells[row, 1];
                    IndiceActivo = objRange.get_Value(Type.Missing).ToString();

                    objRange = (Excel.Range)sheet.Cells[row - 1, 1];
                    string IndiceAnt = objRange.get_Value(Type.Missing).ToString();
                    //i = 1;

                    while (NombreRangos.Contains("IA_" + IndiceActivo))
                    {
                        tienedif = false;

                        dif = Convert.ToInt64(IndiceActivo) - Convert.ToInt64(IndiceAnt);
                        while (dif != 100)
                        {

                            IndiceAnt = "0" + Convert.ToString(Convert.ToInt64(IndiceActivo) - 100);
                            IndiceActivo = IndiceAnt;

                            dif = dif - 100;

                            tienedif = true;
                        }

                        objRange = (Excel.Range)sheet.Cells[row, 1];
                        objRange.Value2 = IndiceAnt;
                       
                        if (tienedif)
                            AddNamedRange(row, 1, "IA_" + Convert.ToString(IndiceAnt));
                        //busco el siguiente activo
                        row++;
                        objRange = (Excel.Range)sheet.Cells[row, 1];
                        IndiceActivo = objRange.get_Value(Type.Missing).ToString();


                    }




                    // busco el numero de fila OTRO para agregarle luego la sumatoria de los indices nuevos
                    try
                    {
                        row = FilaPadre.FirstOrDefault();
                        objRange = (Excel.Range)sheet.Cells[row, 2];
                        string ConceptoPrevio = objRange.get_Value(Type.Missing);
                        if (ConceptoPrevio != null)
                        {
                            ConceptoPrevio = ConceptoPrevio.ToString();
                            if (ConceptoPrevio.Length >= 4)
                            {
                                if (ConceptoPrevio.Substring(0, 4).ToUpper() != "OTRO")
                                {
                                    while (row > 0)
                                    {
                                        objRange = (Excel.Range)sheet.Cells[row, 2];
                                        ConceptoPrevio = objRange.get_Value(Type.Missing);
                                        if (ConceptoPrevio != null)
                                        {
                                            ConceptoPrevio = ConceptoPrevio.ToString();
                                            if (ConceptoPrevio.Length >= 4)
                                            {
                                                if (ConceptoPrevio.Substring(0, 4).ToUpper() == "OTRO")
                                                {

                                                    break;
                                                }
                                            }
                                        }

                                        row--;
                                    }
                                }

                            }
                            else
                            {
                                while (row > 0)
                                {
                                    objRange = (Excel.Range)sheet.Cells[row, 2];
                                    ConceptoPrevio = objRange.get_Value(Type.Missing);
                                    if (ConceptoPrevio != null)
                                    {
                                        ConceptoPrevio = ConceptoPrevio.ToString();
                                        if (ConceptoPrevio.Length >= 4)
                                        {
                                            if (ConceptoPrevio.Substring(0, 4).ToUpper() == "OTRO")
                                            {

                                                break;
                                            }
                                        }
                                    }
                                    row--;
                                }

                            }
                        }
                        else
                        {
                            while (row > 0)
                            {
                                objRange = (Excel.Range)sheet.Cells[row, 2];
                                ConceptoPrevio = objRange.get_Value(Type.Missing);
                                if (ConceptoPrevio != null)
                                {
                                    ConceptoPrevio = ConceptoPrevio.ToString();
                                    if (ConceptoPrevio.Length >= 4)
                                    {
                                        if (ConceptoPrevio.Substring(0, 4).ToUpper() == "OTRO")
                                        {

                                            break;
                                        }
                                    }
                                }
                                row--;
                            }

                        }

                        Excel.Range objRangeJ = ((Excel.Range)sheet.Cells[FilaPadre[0], 1]);
                        objRangeJ.Select();

                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);
                    }

                    try
                    { // limpio si hay error en la formula
                        Excel.Range objRangeI = ((Excel.Range)sheet.Cells[row, 1]).SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors);//obten las celdas con errores

                        //////select all the cells with error formula
                        objRangeI.Clear();
                        objRangeI.Select();
                    }
                    catch (Exception ex)
                    {

                    }



                }



            }
            catch (Exception ex)
            {

              //  MessageBox.Show(ex.Message);

            }



        
        }

        public static void CargaFormulario(string Indice,string Concepto)
        {
                 frmNewExplicaciones NewExplicacion = new frmNewExplicaciones();
                 NewExplicacion.Text = "Explicación índice " + Indice.ToString();
                 if (Concepto != null)
                       NewExplicacion.Text = NewExplicacion.Text + " " + Concepto.ToString();

                  NewExplicacion.ShowDialog();

        }

        private void btnNewExpl_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                Excel.Worksheet ActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
               
                Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell; // Fila activa
                Excel.Range objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row, 9];// voy a la columna 9
                string DebeExplicar = objRange.get_Value(Type.Missing);
                objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row, 1]; // me aseguro que sea la activa en la columna 1
                string Indice = objRange.get_Value(Type.Missing);
                objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row, 2];// voy a la columna 2 de concepto del indice activo
                string Concepto = objRange.get_Value(Type.Missing);
                objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row + 1, 1]; // voy al indice siguiente
                string IndiceSig = objRange.get_Value(Type.Missing);

                if (Indice != null) 
                {
                    if (Indice.ToString().ToUpper().Trim()=="EXPLICACION")
                        MessageBox.Show("El índice " + Indice.ToString()+ "no es válido", "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                    if (DebeExplicar != null)
                    {
                        if (DebeExplicar.ToString().ToUpper() == "SI")
                        {
                            
                            if (IndiceSig != null)
                            {
                                if (IndiceSig.ToString().ToUpper().Trim() == "EXPLICACION")
                                   MessageBox.Show("El índice " + Indice.ToString() + " ya tiene una explicación asociada", "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                else
                                   CargaFormulario(Indice,Concepto);

                                
                            }
                            else
                                CargaFormulario(Indice, Concepto);

                        
                        }
                        else
                          MessageBox.Show("Debe haber una respuesta afirmativa en la columna ' EXPLICAR VARIACION ' para agregar una explicación en el índice " + Indice.ToString(), "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);


                    }
                    else
                         MessageBox.Show("Debe haber una respuesta afirmativa en la columna ' EXPLICAR VARIACION ' para agregar una explicación en el índice " + Indice.ToString(), "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    
                }else
                    MessageBox.Show("El índice no es válido ", "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
            }
            catch { }
        }

        private void btnDelExpl_Click(object sender, RibbonControlEventArgs e)
        {
           
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            string indice = currentCell.Value2;
            if (indice.ToUpper().Trim() == "EXPLICACION")
              currentCell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            else
                MessageBox.Show("La fila seleccionada no es una explicación ", "Eliminar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }
    }
}
