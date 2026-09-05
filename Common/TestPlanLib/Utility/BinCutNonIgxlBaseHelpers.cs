using System.Data;
using System.IO;
using System.Text;

using CommonLib.Extension;

namespace TestPlanLib.Utility
{
    public static class BinCutNonIgxlBaseHelpers
    {
        private const int ConMaxSearchColumn = 10;
        private const int ConMaxSearchRow = 10;

        public static void WriteToFile(string outputFolder, DataTable dataTable, string fileName, bool isCsharp)
        {
            const string lStrEndLine = "End\tend";
            string filePath = Path.Combine(outputFolder, fileName) + ".txt";

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using var sw = new StreamWriter(new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite));
            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                object?[] array = dataTable.Rows[j].ItemArray;
                string lStrContent;
                int i;
                for (i = 0; i < array.Length - 1; i++)
                {
                    lStrContent = array[i]!.ToString()!.Replace("\n", " ") + "\t";
                    sw.Write(lStrContent);
                }
                lStrContent = array[i]!.ToString()!.Replace("\n", " ");
                sw.Write(lStrContent, Encoding.UTF8);
                sw.Write("\r\n");
            }
            if (!isCsharp && !fileName.Contains("bincut_ate_condition_outside") && !fileName.Contains("bincut_ate_condition_eqn_vol") && !fileName.StartsWithIgnoreCase("pwrbinning_") && !fileName.StartsWithIgnoreCase("pwrscreen_"))
            {
                sw.WriteLine(lStrEndLine);
            }
        }

        public static void RemoveColumnsAfterHeader(string startHeader, string endHeader, DataTable dataTable, ref string errMsg)
        {
            bool findEndFlag = false;
            int endColumnNum = dataTable.Columns.Count - 1;
            const string strAllowEqual = "Allow Equal";
            int colAllowEqual = -1;
            int i;
            for (i = 0; i < ConMaxSearchRow; i++)
            {
                int j;
                for (j = 0; j < ConMaxSearchColumn; j++)
                {
                    string lStrCellValue = dataTable.Rows[i][j].ToString()!;
                    if (lStrCellValue.EqualsIgnoreCase(startHeader))
                    {
                        findEndFlag = true;
                        break;
                    }
                }
                if (findEndFlag)
                {
                    break;
                }
            }

            for (int k = 0; k < dataTable.Columns.Count; k++)
            {
                string lStrCellValue = dataTable.Rows[i][k].ToString()!;
                if (lStrCellValue.EqualsIgnoreCase(strAllowEqual))
                {
                    colAllowEqual = k;
                    continue;
                }
                if (lStrCellValue.EqualsIgnoreCase(endHeader))
                {
                    endColumnNum = k;
                }
            }

            if (endColumnNum == dataTable.Columns.Count - 1)
            {
                dataTable.Columns.Add().SetOrdinal(colAllowEqual + 1);
                endColumnNum = colAllowEqual + 1;
                dataTable.Rows[i][endColumnNum] = endHeader;
                errMsg = $"missing {endHeader} column";
            }

            for (int k = dataTable.Columns.Count - 1; k > endColumnNum; k--)
            {
                dataTable.Columns.RemoveAt(k);
            }
        }
    }
}
