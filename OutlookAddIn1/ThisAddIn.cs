using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //CreateMailItem();
        }

        private void CreateMailItem()
        {
            OpenFileDialog attachment = new OpenFileDialog();
            attachment.ShowDialog();
            if (attachment.FileName.Length <= 0)
            {
                return;
            }
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\practice\faculty_ENGR.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string[,] class_details = new string[rowCount + 1, colCount + 1];

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                { 
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        class_details[i, j] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
            }

            //Debug.WriteLine(class_details[1,2]);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Debug.Write(class_details.Length);

            for (int i = 2; i <= rowCount; i++)
            {
                Outlook.MailItem mailItem = (Outlook.MailItem)
                        this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                string course_name = class_details[i,1] + " " + class_details[i,2];
                mailItem.Subject = "Interested in TA opportunity for " + course_name;
                mailItem.To = class_details[i, 4];
                mailItem.Body = "Dear Professor,\n\nI am Pavan Kalyan Thota, enrolled as a full-time graduate student at Indiana University, majoring in Computer Science. After reviewing the " + course_name + " course structure and syllabus, I am interested to be part of the course and hopefully to work with you as TA/AI.\n\nI had completed my bachelor’s degree in computer science and got A + with 105 % in CSCI - B 505 Applied Algorithms course taught by professor Oguzhan Kulecki here in spring 2022.\n\nI am also attaching my resume with this mail. Please let me know if you need any more information. Looking forward to hearing from you.\n\nRegards,\nPavan Thota";
                mailItem.Attachments.Add(attachment.FileName, Outlook.OlAttachmentType.olByValue, 1, "Pavan Resume");
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                //mailItem.Display(false);
                //((Outlook._MailItem)mailItem).Send();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
