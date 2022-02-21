using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
namespace ExcelAddInTest
{
    class JsonUtility
    {
        Workbook wb;

        public JsonUtility(Workbook wb)
        {
            this.wb = wb;
        }

        public string GetObject(Worksheet sheet, Range cell)
        {
            StringBuilder builder = new StringBuilder();
            int row = cell.Row;
            int colMax = sheet.Range["A3"].End[XlDirection.xlToRight].Column;
            builder.Append("{");
            for (int i = 1; i < colMax + 1; i++)
            {
                Range target;
                Range rg = sheet.Cells[row, i];
                string formula = rg.Formula;
                string value = rg.Text;
                string type = sheet.Cells[1, i].Text;
                string field = sheet.Cells[3, i].Text;
                string[] args = Helper.GetArgs(formula);
                builder.Quot(field);
                builder.Append(":");
                Worksheet targetSheet = args != null && args.Length > 1 ? (Worksheet)wb.Worksheets[args[0]] : null;
                switch (type)
                {
                    case "string":
                        builder.Quot(value);
                        break;
                    case "string[]":

                        target = targetSheet.Range[args[1]];
                        builder.Append("[");
                        for (int k = target.Row; k < target.Cells.Count + 1; k++)
                        {
                            string r = targetSheet.Cells[k, target.Column].Text;
                            builder.Quot(r == null ? string.Empty : r);
                            builder.Append(",");
                        }
                        builder.Append("]");
                        break;
                    case "array":
                        target = targetSheet.Range[args[1]];
                        builder.Append("[");
                        for (int k = target.Row; k < target.Cells.Count + 1; k++)
                        {
                            string r = targetSheet.Cells[k, target.Column].Text;
                            builder.Append(r == null ? "0" : r);
                            builder.Append(",");
                        }
                        builder.Append("]");
                        break;
                    case "object":
                        Worksheet sht = targetSheet;
                        target = sht.Range[args[1]];
                        builder.Append(GetObject(sht, target));
                        break;
                    case "object[]":
                        target = targetSheet.Range[args[1]];
                        builder.Append("[");
                        for (int k = target.Row; k < target.Row + target.Cells.Count + 1; k++)
                        {

                            builder.Append(GetObject(targetSheet, targetSheet.Cells[k, 1]));
                            builder.Append(",");
                        }
                        builder.Append("]");
                        break;
                    default:
                        builder.Append(value == null ? "0" : value);
                        break;
                }
                builder.Append(",");
            }
            builder.Append("}");
            return builder.ToString().Replace(",}", "}").Replace(",]", "]");
        }

    }
    static class Helper
    {
        public static string[] GetArgs(string s)
        {
            string[] args = s.Split('!');
            if (args.Length < 2)
            {
                return null;
            }
            args[0] = args[0].Substring(1);
            return args;
        }
        public static void Quot(this StringBuilder bd, string s)
        {
            bd.Append("\"");
            bd.Append(s);
            bd.Append("\"");
        }
    }

}
