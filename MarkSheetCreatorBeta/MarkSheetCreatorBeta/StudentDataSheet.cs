using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
namespace MarkSheetCreator
{
    public class StudentDataSheet
    {
        private static Excel.Application excelAppDataSheet;
        private static Excel.Worksheet publicDataSheetMainSheet;
        private static Excel.Workbook studentDataWorkbookPublic;
        private static List<string> listofDataSheetValues;
        private static List<string> studentDataSheetValues;
        public static List<string> StudentDataSheetValues
        {
            get
            {
                return studentDataSheetValues;
            }
            set
            {
                studentDataSheetValues = value;
            }
        }
        public static List<string> ListofDataSheetValues
        {
            get
            {
                return listofDataSheetValues;
            }
            set
            {
                listofDataSheetValues = value;
            }
        }
        public static Excel.Workbook StudentDataWorkbookPublic
        {
            get
            {
                return studentDataWorkbookPublic;
            }
            set
            {
                studentDataWorkbookPublic = value;
            }
        }
        public static Excel.Worksheet PublicDataSheetMainSheet
        {
            get
            {
                return publicDataSheetMainSheet;
            }
            set
            {
                publicDataSheetMainSheet = value;
            }
        }
        public static Excel.Application ExcelAppDataSheet
        {
            get
            {
                return excelAppDataSheet;
            }
            set
            {
                excelAppDataSheet = value;
            }
        }
        public void ExcelStudentDataSheet()
        {
            studentDataSheetValues = new List<string>();
            listofDataSheetValues = new List<string>();
            Excel.Application excelApp = new Excel.Application();
            excelAppDataSheet = excelApp;
            if (Form1.PublicDataTableFilePath != null)
            {
                Excel.Workbook studentDataWorkbook = excelApp.Workbooks.Open(Form1.PublicDataTableFilePath);
                studentDataWorkbookPublic = studentDataWorkbook;
                Excel.Worksheet dataSheetMainSheet = StudentDataWorkbookPublic.Sheets[1];
                publicDataSheetMainSheet = dataSheetMainSheet;
                int columnCount = publicDataSheetMainSheet.UsedRange.Columns.Count;
                object[,] sheetValues = publicDataSheetMainSheet.UsedRange.Value;
                for (int i = 1; i <= columnCount; i++)
                {
                    if (PublicDataSheetMainSheet.UsedRange.Columns.Value2[1, i] != null)
                    { 
                        listofDataSheetValues.Add(Convert.ToString(PublicDataSheetMainSheet.UsedRange.Columns.Value2[1, i]));
                    }
                }
            }
        }    
    }
}
