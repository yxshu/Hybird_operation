using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace Hybird_operation
{
    class Program
    {
        static void Main(string[] args)
        {
            List<KeyValuePair<string, int>> list = new List<KeyValuePair<string, int>>();
            //maxRange = 100,数据的取值范围 
            //plusMaxRange = 20, 乘除法数据的取值范围
            //TotalFormulaNUM = 95, 总的生成多少道试题
            //CurrentFormulaNUM = 0, 当前生成到第几题
            //columns = 25;每列有多少行
            int maxRange = 100, plusMaxRange = 10, TotalFormulaNUM = 75, CurrentFormulaNUM = 0, rows = 25;
            int answer;
            string formula;
            do
            {
                formula = CreateOuterLayerFormula(maxRange, plusMaxRange, out answer);
                CurrentFormulaNUM++;
                Console.WriteLine(CurrentFormulaNUM + ":" + formula + " = " + answer);
                list.Add(new KeyValuePair<string, int>(formula, answer));
            } while (TotalFormulaNUM - CurrentFormulaNUM > 0);
            if (InsertEXCEL("d:/1.xls", list, rows))
                Console.WriteLine("数据生成完成。");
            Console.ReadLine();
        }


        private static bool InsertEXCEL(string path, List<KeyValuePair<string, int>> list, int rows)
        {
            var workbook = new HSSFWorkbook();
            var table = workbook.CreateSheet("sheet1");
            int columns = (int)Math.Ceiling((double)list.Count / (double)rows);
            for (var i = 0; i < rows; i++)
            {
                var row = table.CreateRow(i);
                for (int j = 0; j < columns; j++)
                {
                    int currentNUM = i * columns + j;
                    if (currentNUM < list.Count)
                    {
                        var cell = row.CreateCell(j);
                        cell.SetCellValue(currentNUM + 1 + ":" + list[currentNUM].Key + "=");
                    }
                    else
                    {
                        break;
                    }
                }
            }
            using (var fs = File.OpenWrite(@path))
            {
                workbook.Write(fs);
            }
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

