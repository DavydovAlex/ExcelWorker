using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Exel_XML.Classes;
using System.Threading;

namespace ExcelWorkerTest
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelWorker test = new ExcelWorker("D:/table.xlsx");
            try
            {
                test.Open(ReadOnly: false);
                int rows = test.Rows;
                int columns = test.Columns;
                //2:51
                for (int i = 1; i <= 10; i++)
                {
                    object[] str = test.ReadRow(i);
                }



                test.Close();        
             }
            catch (Exception e)
            {
                test.Close();
            }
            

        }
    }
}
