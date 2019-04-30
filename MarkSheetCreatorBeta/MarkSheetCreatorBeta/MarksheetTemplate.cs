using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace MarkSheetCreator
{
    public class MarksheetTemplate
    {
        private static Excel.Application excelAppMarksheet;
        private static Excel.Workbook marksheetTemplatePublic;
        private static Excel.Worksheet marksheetMainSheetPublic;

        public static Excel.Worksheet MarksheetMainSheetPublic
        {
            get
            {
                return marksheetMainSheetPublic;
            }
            set
            {
                marksheetMainSheetPublic = value;
            }
        }
        public static Excel.Application ExcelAppMarksheet
        {
            get
            {
                return excelAppMarksheet;
            }
            set
            {
                excelAppMarksheet = value;
            }
        }
        public static Excel.Workbook MarksheetTemplatePublic
        {
            get
            {
                return marksheetTemplatePublic;
            }
            set
            {
                marksheetTemplatePublic = value;
            }
        }
        public void ExcelMarkSheetTemplate()
        {
            Excel.Application excelApp = new Excel.Application();
            excelAppMarksheet = excelApp;
            if (Form1.PublicMarkSheetTemplateFilepath != null)
            {
                Excel.Workbook studentMarkSheetTemplate = excelApp.Workbooks.Open(Form1.PublicMarkSheetTemplateFilepath);
                marksheetTemplatePublic = studentMarkSheetTemplate;
                Excel.Worksheet marksheetMainSheet = MarksheetTemplatePublic.Sheets[1];
                marksheetMainSheetPublic = marksheetMainSheet;
            }
        }
    }
}
