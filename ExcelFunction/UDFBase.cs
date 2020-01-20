using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using System.Security.Permissions;

namespace ExcelFunction
{
    public class UDFBase
    {
        /// <summary>
        /// 解决在某些机器的Excel提示找不到mscoree.dll的问题
        /// 这里在注册表中将该dll的路径注册进去，当使用regasm注册该类库为com组件
        /// 时会调用该方法
        /// </summary>
        /// <param name="type"></param>
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll", RegistryValueKind.String);
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
        }

        private static string GetSubKeyName(Type type, string subKeyName)
        {
            return string.Format("CLSID\\{{{0}}}\\{1}", type.GUID.ToString().ToUpper(), subKeyName);
        }

        /// <summary>
        ///  将Object类的四个公共方法隐藏
        ///  否则将会出现在Excel的UDF函数中
        /// </summary>
        /// <returns></returns>
        [ComVisible(false)]
        public override string ToString()
        {
            return base.ToString();
        }

        [ComVisible(false)]
        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }

        [ComVisible(false)]
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
