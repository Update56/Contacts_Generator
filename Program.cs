using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

internal class Program
{
    private static void Main(string[] args)
    {
        FileInfo fi = new FileInfo("tel.xlsx");
        string path = "Contacts.vcf";
        using (ExcelPackage excelPackage = new ExcelPackage(fi))
        {
            if (!fi.Exists)
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("tel");
                excelPackage.SaveAs(fi);
                return;
            }
            ExcelWorksheet sheet = excelPackage.Workbook.Worksheets[0];
            int i = 1;
            File.WriteAllText(path, "");
            while (sheet.Cells[i, 2].Value != null)
            {
                string tel = sheet.Cells[i, 2].Value.ToString();
                switch (tel.First())
                {
                    case '9':
                        tel += "+7";
                        break;
                    case '8':
                        tel = "+7" + tel.Remove(0, 1);
                        break;
                    case '7':
                        tel += "+";
                        break;

                    default:
                        break;
                }
                File.AppendAllText(path, "BEGIN:VCARD\r\nVERSION:2.1\r\nN:#" + sheet.Cells[i, 1].Value.ToString() + ";Tel;;;\r\nFN:Tel #" + sheet.Cells[i, 1].Value.ToString() + "\r\nTEL;CELL:" + tel + "\r\nEND:VCARD\r\n\r\n");
                i++;
            }
        }
    }
}