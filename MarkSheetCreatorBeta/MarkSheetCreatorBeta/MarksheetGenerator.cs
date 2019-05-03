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
        public void GenerateMarksheets()
        {

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
                        var fileName = DataSheetColumnsToFileName.Aggregate((partialPhrase, word) => $"{partialPhrase}, {word}");
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
}
