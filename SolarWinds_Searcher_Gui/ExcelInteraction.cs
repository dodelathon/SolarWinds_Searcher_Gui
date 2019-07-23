using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using CustomExceptions;
using SolarWinds_Searcher_Gui.Windows.ErrorPopUps;

namespace SolarWinds_Searcher_Gui
{
    public class ExcelInteraction
    {
        private Application excel;
        private _Workbook wb;
        public _Worksheet Feeder;
        private _Worksheet ResultsSheet;

        private SaveFailPopUp sFail;
        private ExcelCloseFailPopUp cFail;
        
        private readonly short SERIALCOL = 1;
        private readonly short EXISTSCOL = 2;
        private readonly short DUPCOL = 3;

        private int col;
        private int row;

        private string sheetName;
        
        public ExcelInteraction(string FilePath, string sName, int _col)
        {
            col = _col;
            row = 2;
            try
            { 
                excel = new Application();
                wb = excel.Workbooks.Open(FilePath);
                sheetName = sName;
                //ReadyResultSheet();
                //excel.Visible = true;   
                SetFeeder();
            }
            catch(SheetNotFoundException e)
            {
                throw new ExcelInterationException(e.Message, e);
            }
            catch(Exception e)
            {
                throw new CantOpenException();
            }

        }

        public ExcelInteraction(string FilePath)
        {
            try
            {
                excel = new Application();
                wb = excel.Workbooks.Open(FilePath);
            }
            catch(Exception e)
            {
                throw new CantOpenException();
            }
            
            //excel.Visible = true;
        }

        public void SetSheetName(string sname)
        {
            sheetName = sname;
            try
            {
                SetFeeder();
            }
            catch (SheetNotFoundException e)
            {
                throw new ExcelInterationException(e.Message, e);
            }
            catch(Exception e)
            {
                throw new CantOpenException();
            }
        }

        private void SetFeeder() 
        {
            try
            {
                Sheets temp = wb.Worksheets;

                foreach (_Worksheet w in temp)
                {
                    if (w.Name.Equals(sheetName))
                    {
                        Feeder = w;
                    }
                }

                if (Feeder == null)
                {
                    throw new SheetNotFoundException(sheetName);
                }
            }
            catch(Exception e)
            {
                throw new CantOpenException();
            }

        }

        private void ReadyResultSheet(string attribute)
        {
            Sheets temp = wb.Worksheets;

            foreach (_Worksheet w in temp)
            {
                if (w.Name.Equals("Results"))
                {
                    ResultsSheet = w;
                }
            }

            if (ResultsSheet == null)
            {
                ResultsSheet = (_Worksheet)wb.Worksheets.Add();
                ResultsSheet.Name = "Results";
            }

            short headerRow = 1;
            ResultsSheet.Cells.Item[headerRow, SERIALCOL].Value2 = attribute;
            ResultsSheet.Cells.Item[headerRow, EXISTSCOL].Value2 = "On Solarwinds?";   
            ResultsSheet.Cells.Item[headerRow, DUPCOL].Value2 = "Duplicates?";

            ResultsSheet.Cells[headerRow, SERIALCOL].Interior.ColorIndex = 6;
            ResultsSheet.Cells[headerRow, EXISTSCOL].Interior.ColorIndex = 6;
            ResultsSheet.Cells[headerRow, DUPCOL].Interior.ColorIndex = 6;

            ResultsSheet.Cells.ColumnWidth = "On Solarwinds?".Length;
        }

        public string AddResult(long row, string sNum, string onSolar, string dups)
        {
            try
            {
                ResultsSheet.Cells.Item[row, SERIALCOL].Value2 = sNum;
                ResultsSheet.Cells.Item[row, EXISTSCOL].Value2 = onSolar;
                ResultsSheet.Cells.Item[row, DUPCOL].Value2 = dups;
                return "Success";
            }
            catch(Exception e)
            {
                return e.Message;
            }
        }

        public string GetNext(long row, long col)
        {
            return Feeder.Cells.Item[row, col].Value2;
        }

        public Collection<string> GetSheets()
        {
            Collection<string> res = null; ;
            try
            {
                res = new Collection<string>();
                Sheets temp = wb.Worksheets;

                foreach (_Worksheet w in temp)
                {
                    res.Add(w.Name);
                }

                return res;
            }
            catch(Exception e)
            {
                throw new CantOpenException();
            }
        }

        public int[] FindStart(int rowsToSearch, int colsToSearch, string searcherAttribute)
        {
            bool found = false;
            int[] retVal = new int[2];
            for (int i = 1; i < rowsToSearch && found == false; i++)
            {
                for(int j = 1; j < colsToSearch && found == false; j++)
                {
                    string temp = (Feeder.Cells.Item[i, j].Value2).ToString();
                    if (temp != null)
                    {
                        Console.WriteLine(temp);
                        temp = temp.ToLower();
                        if (temp.Equals(searcherAttribute.ToLower()))
                        {
                            found = true;
                            retVal[0] = i + 1;
                            retVal[1] = j;
                            col = j;
                            row = i + 1;
                        }
                    }
                }
            }
            return retVal;
        }

        public int GetCount(string attribute)
        {
            int valCounter = 0;
            int conMissCounter = 0;
            row = 2;
            //string val;
            try
            {
                Console.WriteLine(Feeder.Name);
                
                while (conMissCounter < 5)
                {
                    //Range b = Feeder.Cells[row, col];
                    Console.WriteLine(row + " " + col);
                    string val = (string)Feeder.Cells.Item[row, col].Value2;
                    Console.WriteLine(val);
                    if (val == null || val.Equals(""))
                    {
                        conMissCounter++;
                        row++;
                        valCounter++;
                    }
                
                    else
                    {
                        conMissCounter = 0;
                        valCounter++;
                        row++;
                    }
                }
                ReadyResultSheet(attribute);
                return valCounter;
            }
            catch(Exception e)
            {
                return -1;
            }
        }

        public void SetCol(int _col)
        {
            col = _col;
        }

        public void Show()
        {
            excel.Visible = true;
        }

        public void Save()
        {
            try
            {
                if (wb != null)
                {
                    wb.Save();
                }
            }
            catch
            {
                sFail = new SaveFailPopUp
                {
                    Visible = true
                };
            }
        }

        public void Close()
        {
            
            try
            {
                Save();
                wb.Close();
            }
            catch
            {
                cFail = new ExcelCloseFailPopUp
                {
                    Visible = true
                };
            }
        }

        public void DeDupe()
        {
            ResultsSheet.Cells.RemoveDuplicates(1);
        }
    }
}
