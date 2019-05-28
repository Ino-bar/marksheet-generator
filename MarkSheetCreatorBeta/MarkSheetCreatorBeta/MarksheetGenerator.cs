using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace MarkSheetCreator
{
    public class MarksheetGenerator
    {
        List<int> columns = new List<int>();
        List<string> selectedCells = new List<string>();
        public static string fileName;
        public void GenerateMarksheets()
        {

        }
        public void SaveMarksheetsAsNewTab()
        {
            columns = Form1.ListofChosenValues;
            selectedCells = Form1.ListofChosenCells;
            string MarkSheetCompletedfilePath = Form1.PublicDataTableFilePath;
            foreach (Excel.Worksheet sheet in StudentDataSheet.StudentDataWorkbookPublic.Worksheets)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (Excel.Range row in sheet.UsedRange.Rows)
                {
                    if (Form1.NumberOfCheckedCheckboxes != 0)
                    {
                        string[] DataSheetColumnsToFileName = new string[Form1.ListofCheckBoxes.Count];
                        for (int k = 0; k < Form1.ListofCheckBoxes.Count; k++)
                        {
                            if (Form1.ListofCheckBoxes[k].CheckState == CheckState.Checked)
                            {
                                var checkBoxToDataSheet = Form1.ListofCheckBoxes.IndexOf(Form1.ListofCheckBoxes[k]);
                                DataSheetColumnsToFileName[k] = Convert.ToString(sheet.Cells[row.Row, columns[checkBoxToDataSheet] + 1].Value);
                                DataSheetColumnsToFileName[k] = DataSheetColumnsToFileName[k].Replace("/", "-");
                            }
                        }
                        List<string> DataSheetColumnsToFileNameArrayToList = new List<string>(DataSheetColumnsToFileName);
                        DataSheetColumnsToFileNameArrayToList.RemoveAll(item => item == null);

                        fileName = DataSheetColumnsToFileNameArrayToList.Aggregate((partialPhrase, word) => $"{partialPhrase}, {word}");
                        char last = fileName[fileName.Length - 1];
                        if (last.Equals(','))
                        {
                            fileName = fileName.Remove(fileName.Length - 1);
                        }
                        if (fileName.Length > 31)
                        {
                            int charCount = fileName.Length - 31;
                            fileName = fileName.Truncate(31);
                        }
                    }
                    MarksheetTemplate.MarksheetMainSheetPublic.Copy(Type.Missing, MarksheetTemplate.MarksheetTemplatePublic.Sheets[MarksheetTemplate.MarksheetTemplatePublic.Sheets.Count]);
                    MarksheetTemplate.MarksheetTemplatePublic.Sheets[MarksheetTemplate.MarksheetTemplatePublic.Sheets.Count].Name = fileName;
                    int i = 0;
                    if (i <= selectedCells.Count)
                    {
                        foreach (var listItem in selectedCells)
                        {
                            MarksheetTemplate.MarksheetTemplatePublic.Application.DisplayAlerts = false;
                            MarksheetTemplate.MarksheetTemplatePublic.Sheets[MarksheetTemplate.MarksheetTemplatePublic.Sheets.Count].Range[selectedCells[i]].Value = Convert.ToString(sheet.Cells[row.Row, columns[i] + 1].Value);
                            i++;
                        }
                    }
                }
                MarksheetTemplate.MarksheetTemplatePublic.SaveAs(Form1.PublicCompletedMarksheetSaveLocation + "\\" + Convert.ToString(fileName), FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
        }
        public void GetChosenColumnsAndData()
        {
            columns = Form1.ListofChosenValues;
            selectedCells = Form1.ListofChosenCells;
            string MarkSheetCompletedfilePath = Form1.PublicDataTableFilePath;
            foreach (Excel.Worksheet sheet in StudentDataSheet.StudentDataWorkbookPublic.Worksheets)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (Excel.Range row in sheet.UsedRange.Rows)
                {
                    int i = 0;
                    if (i <= selectedCells.Count)
                    {
                        foreach (var listItem in selectedCells)
                        { 
                            MarksheetTemplate.MarksheetTemplatePublic.Application.DisplayAlerts = false;
                            MarksheetTemplate.MarksheetMainSheetPublic.Range[selectedCells[i]].Value = Convert.ToString(sheet.Cells[row.Row, columns[i] + 1].Value);
                            i++;
                        }
                    }
                    if (Form1.NumberOfCheckedCheckboxes != 0)
                    {
                        string[] DataSheetColumnsToFileName = new string[Form1.ListofCheckBoxes.Count];
                        for (int k = 0; k < Form1.ListofCheckBoxes.Count; k++)
                        { 
                            if (Form1.ListofCheckBoxes[k].CheckState == CheckState.Checked)
                            {
                                var checkBoxToDataSheet = Form1.ListofCheckBoxes.IndexOf(Form1.ListofCheckBoxes[k]);
                                DataSheetColumnsToFileName[k] = Convert.ToString(sheet.Cells[row.Row, columns[checkBoxToDataSheet] + 1].Value);
                                DataSheetColumnsToFileName[k] = DataSheetColumnsToFileName[k].Replace("/", "-");
                            }
                        }
                        List<string> DataSheetColumnsToFileNameArrayToList = new List<string>(DataSheetColumnsToFileName);
                        DataSheetColumnsToFileNameArrayToList.RemoveAll(item => item == null);

                        var fileName = DataSheetColumnsToFileNameArrayToList.Aggregate((partialPhrase, word) => $"{partialPhrase}, {word}");
                        char last = fileName[fileName.Length - 1];
                        if (last.Equals(','))
                        {
                            fileName = fileName.Remove(fileName.Length - 1);
                        }
                        MarksheetTemplate.MarksheetTemplatePublic.SaveAs(Form1.PublicCompletedMarksheetSaveLocation + "\\" + Convert.ToString(fileName), FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                    //MarksheetTemplate.MarksheetTemplatePublic.SaveAs(Form1.PublicCompletedMarksheetSaveLocation + "\\" + Convert.ToString(sheet.Cells[row.Row, columns[0] + 1].Value) + ", " + Convert.ToString(sheet.Cells[row.Row, columns[1] + 1].Value), FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                }
            }
            System.Windows.Forms.MessageBox.Show("Forms Complete");
        }
        public void CloseExcel()
        {
            StudentDataSheet.StudentDataWorkbookPublic.Close(0);
            StudentDataSheet.ExcelAppDataSheet.Quit();
            MarksheetTemplate.MarksheetTemplatePublic.Close(0);
            MarksheetTemplate.ExcelAppMarksheet.Quit();
        }
    }
    public static class StringExt
    {
        public static string Truncate(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }
}
