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
            WpsReg.wpsRegistry();//wpsע��
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
            WpsReg.wpsUnRegistry();//wps����ע��
            GC.Collect();//��������

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

    public class sthExcelFunction//excel��ͨ��ʽ
    {
     
        [ExcelFunction(ExplicitRegistration = true, Category = "�ַ�������", Description = "�������Ӳ��������ַ���")]
        public static string sthLink(
            [ExcelArgument(Name = "����", Description = "���ӷ���")] string symbol,
            [ExcelArgument(Name = "����", Description = "������Ҫ���ӵ�����")] params object[] optional)
        {
            StringBuilder str = new StringBuilder(100);
            foreach (var arg in optional)
            {
                if (arg is object[,])
                {
                    int rows = ((object[,])arg).GetLength(0);//������
                    int cols = ((object[,])arg).GetLength(1);//������
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

        [ExcelFunction(ExplicitRegistration = true, Category = "�滻�������", Description = "�滻�ַ����ж������")]
        public static string sthReplace(
            [ExcelArgument(Name = "����", Description = "�������")] string symbol,
            [ExcelArgument(Name = "����", Description = "������Ҫ�滻������ŵ��ַ���")] string str)
        {
            str = Regex.Replace(str, symbol + "+", symbol);//�滻�ظ���
            str = Regex.Replace(str, "^" + symbol + "+|" + symbol + "+$", "");//�滻��β
            return str;
        }

        [ExcelFunction(ExplicitRegistration = true, Category = "�ֽ�", Description = "���طֽ���n������")]
        public static string sthFenJie(
            [ExcelArgument(Name = "����", Description = "�����ַ���")] string str,
            [ExcelArgument(Name = "����", Description = "�ָ����")] string symbol,
            [ExcelArgument(Name = "����", Description = "���طֽ���n������")] int number)
        {
            return str.Split(symbol)[number - 1];
        }
    }

}