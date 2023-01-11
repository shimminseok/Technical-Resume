using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
namespace LearnDataTable
{
    internal class ExcelReader
    {
        // 시트 레코드 칼럼
        Dictionary<int, Dictionary<int, List<string>>> excelDatas;

        // 시트이름 , 인덱스 번호 , 컬럼이름,셀값
        public Dictionary<string, Dictionary<string, Dictionary<string,string>>> converDatas;
        
        List<string> _sheetNames;

        public ExcelReader()
        {
            excelDatas = new Dictionary<int, Dictionary<int, List<string>>>();
            _sheetNames = new List<string>();
            converDatas = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>(); 
        }
        #region[외부 사용 함수]
        public bool ReadExcel(string fullPath)
        {
            fullPath =Directory.GetCurrentDirectory() +"\\"+ fullPath;
            object misValue = System.Reflection.Missing.Value;

            Excel.Application oXL = new Excel.Application();
            Excel.Workbooks oWBooks = (Excel.Workbooks)oXL.Workbooks;
            Excel.Workbook oWB = oWBooks.Open(fullPath, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue
                , misValue, misValue, misValue);
            Excel.Sheets oSheets = oWB.Sheets;
            SaveDictonary(oSheets);
            oXL.Visible = false;
            oXL.UserControl = true;
            oXL.DisplayAlerts = false;
            oXL.Quit();

            ReleaseExcelObject(oSheets);
            ReleaseExcelObject(oWB);
            ReleaseExcelObject(oWBooks);
            ReleaseExcelObject(oXL);
            SaveExcelData();
            return true;
        }
        #endregion


        #region[내부 사용 함수]
        void ReleaseExcelObject(object obj)
        {
            try
            {
                if(obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj= null;
                }
            }
            catch(Exception excep)
            {
                obj= null;
                Console.WriteLine(excep.Message);
            }
            finally
            {
                GC.Collect();
            }
        }
        int ExcelColumCount(Excel.Worksheet oSheet,Excel.Range oRng)
        {
            int colCount = oRng.Column ;
            for(int i = 1; i <= colCount; i++)
            {
                Excel.Range cell = (Excel.Range)oSheet.Cells[1,i];
                if(cell.Value == null)
                {
                    ReleaseExcelObject(cell);
                    Console.WriteLine("{0} sheet에 비어 있는 셀이 존재 합니다.", oSheet.Name);
                    colCount = i - 1;
                    break;
                }
            }

            return colCount;
        }
        List<string> ExcelColumsName(int length)
        {
            List<string> columsList = new List<string>();

            int basenum = 26;
            for (int i = 0; i < length; i++)
            {
                if (i / basenum == 0)
                    columsList.Add(Convert.ToString((char)(65 + i)));
                else
                {
                    string name = Convert.ToString((char)(64 + (i / basenum)));
                    name += Convert.ToString((char)(65+(i% basenum)));
                    columsList[i] = name;
                }
            }
            return columsList;
        }
        void SaveDictonary(Excel.Sheets oSheets)
        {
            for(int i = 1; i <= oSheets.Count; i++)
            {
                List<string> colums;
                Excel.Worksheet oSheet = (Excel.Worksheet)oSheets.get_Item(i);
                _sheetNames.Add(oSheet.Name);
                Excel.Range oRng = oSheet.get_Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int colCount = ExcelColumCount(oSheet, oRng);
                colums = ExcelColumsName(colCount);
                //sheet정보를 담을 장소.
                Dictionary<int,List<string>> sheetdata = new Dictionary<int,List<string>>();
                for(int col =1; col<= colums.Count; col++)
                {
                    int count = 0;
                    List<string> columData = new List<string>();

                    Excel.Range collCell = (Excel.Range)oSheet.Columns[col];
                    Excel.Range range = oSheet.get_Range(colums[col - 1] + "1", collCell);

                    foreach(object cell in range.Value)
                    {
                        if (count < oRng.Row)
                        {
                            count++;
                            if (cell == null)
                                columData.Add("");
                            else
                                columData.Add(cell.ToString());
                        }
                        else
                            break;
                    }
                    sheetdata.Add(col, columData);
                    ReleaseExcelObject(range);
                    ReleaseExcelObject(collCell);
                   
            
                }
                excelDatas.Add(i, sheetdata);
            }

        }

        public void PrintExcel()
        {
            foreach (/*Dictionary<int,List<string>> */ var sheetdata in excelDatas.Values) // 시트 
            {
                {
                    foreach (/*List<string>*/ var data in sheetdata.Values)   
                    {
                        for (int j = 0; j < data.Count; j++)  
                        {

                            Console.Write(data[j] + "\t");
                        }
                        Console.WriteLine();
                    }
                }
            } 

            
        }

        public void ShowExcelOrigin()
        {
            int count = 0;
            //foreach(int id in excelDatas.Keys)
            //{   
            //    Console.WriteLine("====[{0}]========================================================================",_sheetNames[count++]);
            //    Dictionary<int,List<string>> sheet = excelDatas[id];
            //    foreach(int key in sheet.Keys)
            //    {
            //        List<string> ColumnData = sheet[key];
            //        for(int i=0; i<ColumnData.Count; i++)
            //        {
            //            Console.WriteLine(ColumnData[i]);
            //            //if (i < ColumnData.Count - 1)
            //                //Console.Write("\t");
            //        }
            //        Console.WriteLine();
            //    }
            //}
            foreach (int id in excelDatas.Keys)
            {
                Dictionary<int, List<string>> sheet = excelDatas[id];
                int key =0;
                foreach(int keyV in sheet.Keys)
                {
                    key = keyV;
                    break;
                }
                if(key == 0)
                {
                    return;
                }
               
                for(int i=0; i<sheet[key].Count; i++)
                {
                    foreach(int keyV in sheet.Keys)
                    {
                        Console.Write(sheet[keyV][i]);
                        if(i< sheet[key].Count )
                        {
                            Console.Write("\t");
                        }
                    }
                    Console.WriteLine();
                }
            }
        }

       void SaveExcelData()
        {
            for (int i=1; i<= excelDatas.Count; i++) // 시트 수 
            { 
                //시트 데이터 저장
                Dictionary<string, Dictionary<string, string>> cellData = new Dictionary<string, Dictionary<string, string>>();
                for (int j = 1; j < excelDatas[i][1].Count; j++) // 컬럼 수
                {
                    string index = null; // 인덱스 이름
                    Dictionary<string, string> columndatas = new Dictionary<string, string>(); //인덱스와 값 저장
                    for (int k = 1; k <= excelDatas[i].Count; k++) //필드 수
                    {
                        index = excelDatas[i][1][j]; // 인덱스값
                        string CName = excelDatas[i][k][0].ToString(); //  컬럼이름
                        
                        columndatas.Add(CName, excelDatas[i][k][j].ToString()); //컬럼이름 , 값
                        // 배열[시트][가로][세로]
                    }
                    cellData.Add(index, columndatas); // 인덱스번호 , 데이터
                }
                converDatas.Add(_sheetNames[i-1], cellData); // 시트이름 , 시트데이터
            }
        }
        #endregion

        public void ShowCon()
        {
            SaveExcelData();
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, string>>> sheetdata in converDatas)
            {
                Console.WriteLine(sheetdata.Key);
                bool s = true;
                foreach (var data in sheetdata.Value.Values)
                {
                    if (s)
                    {
                        foreach (var d in data)
                        {
                            Console.Write(d.Key + "\t");
                        }
                        s = false;
                    }
                    Console.WriteLine();
                    foreach (var d in data)
                    {
                        Console.Write(d.Value +"\t");
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
            }
        }
    }
}
