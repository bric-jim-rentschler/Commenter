using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Reflection;


namespace Commenter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            AddCommentsToExcel();
            Application.Run(new Form1());
        }
        public static void AddCommentsToExcel()
        {
            Excel.Application excelApp = new Excel.ApplicationClass();
            string myPath = @"C:\Immigration\fy2015_table1.xls";

            excelApp.Workbooks.Open(myPath, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)excelApp.Application.ActiveWorkbook.Sheets[1]);

            Excel.Range allCells = (Excel.Range)activeWorksheet.UsedRange;

            foreach (Excel.Range cell in allCells){
                if (cell.Value != null && !cell.Value.Equals(" "))
                {
                    cell.AddComment(cell.Address.ToString() + " value=" + cell.Value);
                }
                
            }
            excelApp.ActiveWorkbook.Save();
            excelApp.Quit();
            Console.WriteLine("Done comments");
        }

    }
}
