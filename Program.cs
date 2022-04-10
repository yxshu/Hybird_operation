

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace Hybird_operation
{
    class Program
    {
        private static Stopwatch watch = new Stopwatch();
        static void Main(string[] args)
        {
            watch.Start();
            //maxRange = 100,数据的取值范围 
            //plusMaxRange = 20, 乘除法数据的取值范围
            //TotalFormulaNUM = 95, 总的生成多少道试题
            //columns = 25;每列有多少行
            //sheetNUM = 20;共生成多少份
            int maxRange = 100, plusMaxRange = 20, TotalFormulaNUM = 66, columns = 3, sheetNUM = 20;
            string  path = "C://Users//admin//Desktop//1.xlsx", title = "日期______ 开始时间______  结束时间______";
            if (InsertExcels(path, title, createDATA(maxRange, plusMaxRange, TotalFormulaNUM, columns, sheetNUM), columns))
                Console.WriteLine(string.Format("{0} Articles Formula Generation Success。", TotalFormulaNUM * sheetNUM));
            watch.Stop();
            Console.WriteLine("Total time is {0} MS.", watch.ElapsedMilliseconds);
            Console.ReadLine();
        }
        /// <summary>
        /// 按要求生成数据
        /// </summary>
        /// <param name="maxRange">生成算式的最大值</param>
        /// <param name="plusMaxRange">乘除法算式中最大的值</param>
        /// <param name="TotalFormulaNUM">每页生成算式的数量</param>
        /// <param name="columns">每页算式的列数</param>
        /// <param name="sheetNUM">生成的页数</param>
        /// <returns></returns>
        private static List<List<KeyValuePair<string, int>>> createDATA(int maxRange, int plusMaxRange, int TotalFormulaNUM, int columns, int sheetNUM)
        {
            List<List<KeyValuePair<string, int>>> list = new List<List<KeyValuePair<string, int>>>();
            for (int i = 0; i < sheetNUM; i++)
            {
                List<KeyValuePair<string, int>> sheetdata = new List<KeyValuePair<string, int>>();
                int CurrentFormulaNUM = 0;
                do
                {
                    string formula = CreateOuterLayerFormula(maxRange, plusMaxRange, out int answer);
                    CurrentFormulaNUM++;
                    sheetdata.Add(new KeyValuePair<string, int>(formula, answer));
                } while (TotalFormulaNUM - CurrentFormulaNUM > 0);
                list.Add(sheetdata);
            }
            Console.WriteLine(string.Format("Time to generate data is {0} MS. ", watch.ElapsedMilliseconds));
            return list;
        }

        private static bool InsertExcels(string path, string title, List<List<KeyValuePair<string, int>>> list, int columns)
        {
            var workbook = new XSSFWorkbook();
            for (int i = 0; i < list.Count; i++)
            {
                var sheet = workbook.CreateSheet("sheet" + (i + 1));
                sheet.PrintSetup.Scale = 100;
                sheet.PrintSetup.PaperSize = 9;
                int rows = (int)Math.Ceiling((double)list[i].Count / (double)columns);//计算总的包含多少行数据

                sheet.DefaultColumnWidth = 27;//设置默认的列宽
                                              //sheet1.DefaultRowHeight = 25 * 20;//设置默认行高

                var TitleRow = sheet.CreateRow(0);//新建标题行
                TitleRow.HeightInPoints = 60;//设置标题行的行高；
                var Titlecell = TitleRow.CreateCell(0);
                Titlecell.SetCellValue(title);
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, columns - 1));//设置标题行合并居中

                ICellStyle Titlestyle = workbook.CreateCellStyle();//其他2个样式1：用于标题行
                ICellStyle cellStyle = workbook.CreateCellStyle();//样式2：用于其他行
                Titlestyle.Alignment = HorizontalAlignment.Center;
                Titlestyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.Alignment = HorizontalAlignment.Justify;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.BorderTop = BorderStyle.Thin;
                cellStyle.BorderRight = BorderStyle.Thin;
                cellStyle.BorderBottom = BorderStyle.Thin;
                cellStyle.BorderLeft = BorderStyle.Thin;

                IFont Titlefont = workbook.CreateFont();
                IFont font = workbook.CreateFont();
                Titlefont.FontHeight = 20 * 20;
                font.FontHeight = 16 * 20;
                Titlefont.FontName = "楷体";
                font.FontName = "宋体";
                Titlestyle.SetFont(Titlefont);
                cellStyle.SetFont(font);

                Titlecell.CellStyle = Titlestyle;//设置标题行样式

                for (var j = 0; j < rows + 1; j++)
                {
                    var row = sheet.CreateRow(j + 1);//因为第一行是标题行，所以这里从1开始
                    row.HeightInPoints = 30;
                    for (int k = 0; k < columns; k++)
                    {
                        int currentNUM = j * columns + k;
                        if (currentNUM < list[i].Count)
                        {
                            var cell = row.CreateCell(k);
                            cell.SetCellValue(list[i][currentNUM].Key + "=");
                            cell.CellStyle = cellStyle;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            Console.WriteLine("Time to generate EXCEL table is {0} MS.", watch.ElapsedMilliseconds);
            using (var fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                workbook.Write(fs);
            }
            Console.WriteLine("Time to insert into hard-disk is {0} MS.", watch.ElapsedMilliseconds);
            return true;
        }

        private static string CreateOuterLayerFormula(int maxRange, int plusMaxRange, out int answer)
        {
            Random rand = new Random();
            int factora, factorb;
            string[] opera = { "+", "-", "×", "÷" };
            int innerSubmark, submark = rand.Next(opera.Length);
            string InnerLayerFormula, Formula;
            bool isInvers = false;
            do
            {
                factora = rand.Next(maxRange + 1);
                InnerLayerFormula = CreateFormula(maxRange, plusMaxRange, out factorb, out innerSubmark);
            } while (isLoop(factora, factorb, opera[submark], plusMaxRange, out isInvers));
            switch (submark)
            {
                case 0:
                    answer = factora + factorb;
                    if (innerSubmark < 2)
                        Formula = factora + opera[submark] + "(" + InnerLayerFormula + ")";
                    else
                    {
                        Formula = factora + opera[submark] + InnerLayerFormula;
                    }
                    break;
                case 1:
                    if (isInvers == true)
                    {
                        answer = factorb - factora;
                        Formula = InnerLayerFormula + opera[submark] + factora;
                    }
                    else
                    {
                        answer = factora - factorb;
                        if (innerSubmark < 2) Formula = factora + opera[submark] + "(" + InnerLayerFormula + ")";
                        else
                        {
                            Formula = factora + opera[submark] + InnerLayerFormula;
                        }
                    }
                    break;
                case 2:
                    answer = factora * factorb;
                    if (innerSubmark != 2)
                        Formula = factora + opera[submark] + "(" + InnerLayerFormula + ")";
                    else
                    {
                        Formula = factora + opera[submark] + InnerLayerFormula;
                    }
                    break;
                case 3:
                    if (isInvers == true)
                    {
                        answer = factorb / factora;
                        if (innerSubmark < 2)
                        {
                            Formula = "(" + InnerLayerFormula + ")" + opera[submark] + factora;
                        }
                        else { Formula = InnerLayerFormula + opera[submark] + factora; }
                    }
                    else
                    {
                        answer = factora / factorb;
                        if (innerSubmark != 2)
                            Formula = factora + opera[submark] + "(" + InnerLayerFormula + ")";
                        else
                        {
                            Formula = factora + opera[submark] + InnerLayerFormula;
                        }
                    }
                    break;
                default: answer = 0; Formula = String.Empty; break;
            }
            return Formula;
        }
        /// <summary>
        /// 生成一个算式，
        /// 包括：一、减法验证大数减小数；
        /// 二、除法验证能除尽
        /// 三、防出现0和1太简单的题目
        /// </summary>
        /// <param name="maxRange">数值的最大范围</param>
        /// <param name="answer">答案</param>
        /// <returns>算式</returns>
        private static string CreateFormula(int maxRange, int plusMaxRange, out int answer, out int submark)
        {
            Random rand = new Random();
            int factora, factorb;
            string[] opera = { "+", "-", "×", "÷" };
            submark = rand.Next(opera.Length);
            bool isInvers = false;
            do
            {
                factora = rand.Next(maxRange + 1);
                factorb = rand.Next(maxRange + 1);
            } while (isLoop(factora, factorb, opera[submark], plusMaxRange, out isInvers));
            if (isInvers == true)
            {
                int temp = factora;
                factora = factorb;
                factorb = temp;
            }
            switch (submark)
            {
                case 0: answer = factora + factorb; break;
                case 1: answer = factora - factorb; break;
                case 2: answer = factora * factorb; break;
                case 3: answer = factora / factorb; break;
                default: answer = 0; break;
            }
            return factora + opera[submark] + factorb;
        }
        private static bool isLoop(int factora, int factorb, string opera, int plusMaxRange, out bool isverse)
        {
            isverse = false;
            bool result = false;
            switch (opera)
            {
                case "+": if (Math.Min(factora, factorb) < 2) result = true; break;
                case "-": if (Math.Min(factora, factorb) < 2) result = true; else if (factorb > factora) isverse = true; break;
                case "×": if (factora > plusMaxRange || factorb > plusMaxRange || Math.Min(factora, factorb) < 2) result = true; break;
                case "÷":
                    if (factora > plusMaxRange || factorb > plusMaxRange || factora == factorb || Math.Min(factora, factorb) < 2) result = true;
                    else
                    {
                        float remainder;
                        if (factora >= factorb) remainder = factora % factorb;
                        else { remainder = factorb % factora; isverse = true; }
                        if (remainder != 0) result = true;
                    }
                    break;
                default:
                    break;
            };
            return result;
        }
    }
}

