using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace TaskRPA.Writer
{
    /// <summary>
    /// Class that allows to write a list in excel document.
    /// </summary>
    public class ExcelWriter : IWriter<Microwave>
    {
        private Application application;
        private Workbook workBook;

        public string Path { get; set; }

        public ExcelWriter(string path)
        {
            Path = path;
        }

        public void Write(List<Microwave> list)
        {
            if (list == null)
            {
                throw new ArgumentNullException("List cannot be null");
            }

            try
            {
                application = new Application();
                workBook = application.Workbooks.Add(Type.Missing);
                application.SheetsInNewWorkbook = 1;
                application.DisplayAlerts = false;
                Worksheet sheet = (Worksheet)application.Worksheets.get_Item(1);
                sheet.Name = "Microwaves";

                for (int i = 1; i <= list.Count; i++)
                {
                    sheet.Cells[i, 1] = list[i - 1].Title;
                    sheet.Cells[i, 2] = list[i - 1].Price;
                    sheet.Cells[i, 3] = list[i - 1].Href;
                }

                workBook.SaveAs(Path);
            }
            finally
            {
                workBook.Close();
                application.Quit();
            }
        }
    }
}
