using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using LitJson;

namespace LearnDataTable
{

    internal class CreateDatFilecs
    {
        #region[TextFile Write]
        //public void Writer(string fullpath)
        //{
        //    ExcelReader excel = new ExcelReader();
        //    excel.ReadExcel(fullpath);

        //    StreamWriter sw;
        //    Dictionary<string, Dictionary<string, Dictionary<string, string>>> data = excel.converDatas;

        //    foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, string>>> sheet in data)
        //    {
        //        sw = new StreamWriter(sheet.Key + ".dat");
        //        bool s = true;
        //        foreach (var sheetdata in sheet.Value.Values)
        //        {
        //            if (s)
        //            {
        //                foreach (var d in sheetdata)
        //                {
        //                    sw.Write(d.Key + "\t");
        //                }
        //                s = false;
        //            }
        //            sw.WriteLine();
        //            foreach (var d in sheetdata)
        //            {
        //                sw.Write(d.Value + "\t");
        //            }
        //            sw.WriteLine();
        //        }
        //        sw.Close();
        //    }

        //}

        //public void Writer2(Dictionary<string, Dictionary<string, Dictionary<string, string>>> data)
        //{

        //    StreamWriter sw;


        //    foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, string>>> sheet in data)
        //    {
        //        sw = new StreamWriter(sheet.Key + ".dat");
        //        bool s = true;
        //        foreach (var sheetdata in sheet.Value.Values)
        //        {
        //            if (s)
        //            {
        //                foreach (var d in sheetdata)
        //                {
        //                    //sw.Write( d.Key + "\t");
        //                    sw.Write(string.Format("{0} \t", d.Key).PadRight(d.Key.Length));
        //                }
        //                s = false;
        //            }
        //            sw.WriteLine();
        //            foreach (var d in sheetdata)
        //            {
        //                //sw.Write(d.Value + "\t");
        //                sw.Write(string.Format("{0} \t\t", d.Value).PadRight(d.Value.Length));

        //            }
        //            sw.WriteLine();
        //        }
        //        sw.Close();
        //    }

        //}
        #endregion[TextFile Write]

        #region[JsonFile Write]
        public static void ExcelDataToJsonFile(Dictionary<string, Dictionary<string, Dictionary<string, string>>> excelData)
        {
            //foreach(string key in excelData.Keys)
            //{
            //    Dictionary<string, Dictionary<string, string>> sheet = excelData[key];
            //    string fileName = Directory.GetCurrentDirectory() +"\\" +key + ".Json";
            //    StreamWriter sw = new StreamWriter(fileName, false, Encoding.Unicode);
            //    WriteJsonFileFromSheet(sw, key, sheet);
            //    sw.Close();
            //}

            // KeyValuePair 해시테이블의 노드값을 가져온다
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, string>>> sheetData in excelData)
            {
                string fileName = Directory.GetCurrentDirectory() + "\\" + sheetData.Key + ".Json";
                StreamWriter sw = new StreamWriter(fileName, false, Encoding.Unicode);
                WriteJsonFileFromSheet(sw, sheetData.Key, sheetData.Value);
                sw.Close();
            }
        }
        public static void WriteJsonFileFromSheet(StreamWriter sw,string sheetName,Dictionary<string, Dictionary<string, string>> sheet)
        {
            if(sw != null)
            {
                //컬럼 이름을 담아둔다.
                List<string> columnList = new List<string>();
                foreach(string idx in sheet.Keys)
                {
                    Dictionary<string, string> record = sheet[idx];
                    foreach (string column in record.Keys)
                    {
                        columnList.Add(column);
                    }
                    break;
                }
                //Json파일을 만들어서 쓴다.
                JsonWriter writer = new JsonWriter(sw);
                //Object는 중괄호이다. Array는 대괄호
                writer.WriteObjectStart();
                writer.WritePropertyName(sheetName);
                writer.WriteArrayStart();
                foreach (Dictionary<string,string> record in sheet.Values)
                {
                    writer.WriteObjectStart();
                    foreach(string key in record.Keys)
                    {
                        writer.WritePropertyName(key);
                        writer.Write(record[key]);
                    }
                    writer.WriteObjectEnd();
                }
                writer.WriteArrayEnd();
                writer.WriteObjectEnd();

            }
        }
        #endregion[JsonFile Write]
    }
}
