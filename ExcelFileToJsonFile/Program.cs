using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LearnDataTable
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string fullPath = "ExcelData";
            ExcelReader ExReader = new ExcelReader();
            CreateDatFilecs CF = new CreateDatFilecs();
            if (!ExReader.ReadExcel(fullPath))
            {
                Console.WriteLine("{0} 파일 로드 실패", fullPath);
            }
            else
            {
                CreateDatFilecs.ExcelDataToJsonFile(ExReader.converDatas);
            }
            //ExReader.PrintExcel();
            //ExReader.ShowCon();

            //CF.Writer(fullPath);
            //CF.Writer2(ExReader.converDatas);
        }
    }
}
