using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using System;
using System.Collections.Concurrent;
using System.Linq;
using WPSExcelDna;
//using System.Numerics;
using System.Text;
using System.Text.RegularExpressions;
//using Microsoft.Office.Interop.Excel;
//using static System.Net.Mime.MediaTypeNames;
//using NumSharp;
//using static JsonRead.GlobalDict;

namespace WpsFunction
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            RegisterFunctions();
            WpsReg.wpsRegistry();//wps注册
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
            WpsReg.wpsUnRegistry();//wps反向注册
            GC.Collect();//垃圾回收

        }

        public void RegisterFunctions()
        {
            // There are various options for wrapping and transforming your functions
            // See the Source\Samples\Registration.Sample project for a comprehensive example
            // Here we just change the attribute before registering the functions
            ExcelRegistration.GetExcelFunctions()
                             .Select(UpdateHelpTopic)
                             .ProcessParamsRegistrations()
                             .RegisterFunctions();

        }

        public ExcelFunctionRegistration UpdateHelpTopic(ExcelFunctionRegistration funcReg)
        {
            funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
            return funcReg;
        }
    }

    public class sthExcelFunction//excel普通公式
    {
     
        [ExcelFunction(ExplicitRegistration = true, Category = "字符串连接", Description = "根据连接参数连接字符串")]
        public static string sthLink(
            [ExcelArgument(Name = "符号", Description = "连接符号")] string symbol,
            [ExcelArgument(Name = "数据", Description = "输入需要连接的数据")] params object[] optional)
        {
            StringBuilder str = new StringBuilder(100);
            foreach (var arg in optional)
            {
                if (arg is object[,])
                {
                    int rows = ((object[,])arg).GetLength(0);//行数量
                    int cols = ((object[,])arg).GetLength(1);//列数量
                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < cols; j++)
                        {
                            if (((object[,])arg)[i, j] is not ExcelEmpty) { str.Append(((object[,])arg)[i, j] + symbol); }
                        }
                    }
                }
                else
                {
                    if (arg is not ExcelEmpty) { str.Append(arg + symbol); }//ExcelEmpty
                }
            }
            return sthReplace(symbol, str.ToString());//string.Join(symbol, strArray);
        }

        [ExcelFunction(ExplicitRegistration = true, Category = "替换多余符号", Description = "替换字符串中多余符号")]
        public static string sthReplace(
            [ExcelArgument(Name = "符号", Description = "多余符号")] string symbol,
            [ExcelArgument(Name = "数据", Description = "输入需要替换多余符号的字符串")] string str)
        {
            str = Regex.Replace(str, symbol + "+", symbol);//替换重复项
            str = Regex.Replace(str, "^" + symbol + "+|" + symbol + "+$", "");//替换首尾
            return str;
        }

        [ExcelFunction(ExplicitRegistration = true, Category = "分解", Description = "返回分解后第n个数据")]
        public static string sthFenJie(
            [ExcelArgument(Name = "数据", Description = "输入字符串")] string str,
            [ExcelArgument(Name = "符号", Description = "分割符号")] string symbol,
            [ExcelArgument(Name = "索引", Description = "返回分解后第n个数据")] int number)
        {
            return str.Split(symbol)[number - 1];
        }
    }

}