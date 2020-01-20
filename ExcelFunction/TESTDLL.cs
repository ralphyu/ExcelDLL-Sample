using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ExcelFunction
{
    [Guid("5CBFBF06-2949-4DC1-8FF6-87B6864E9FCB")]
    [ClassInterface(ClassInterfaceType.AutoDual), ComVisible(true)]
    public class TESTDLL : UDFBase
    {
        //测试函数1
        public int FuncTest(int a)
        {
            try
            {
                return a * 10;
            }
            catch (Exception )
            {
                return -999;
            }
        }

        //测试函数2
        public double FuncASSD(int a,double b,double c)
        {
            return FuncTest(a) + b * c;
        }
    }
}
