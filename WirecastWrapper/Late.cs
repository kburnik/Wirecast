using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace WirecastWrapper
{
    class Late
    {
        
        public static void Set(object obj, string sProperty, object oValue)
        {
            object[] oParam = new object[1];
            oParam[0] = oValue;
            obj.GetType().InvokeMember(sProperty, BindingFlags.SetProperty, null, obj, oParam);
        }

         
        public static object Get(object obj, string sProperty, object oValue)
        {
            object[] oParam = new object[1];
            oParam[0] = oValue;
            return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oParam);
        }

         
        public static object Get(object obj, string sProperty, object[] oValue)
        {
            return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oValue);
        }

         
        public static object Get(object obj, string sProperty)
        {
            return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, null);
        }

         
        public static object Invoke(object obj, string sProperty, object[] oParam)
        {
            return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
        }

         
        public static object Invoke(object obj, string sProperty, object oValue)
        {
            object[] oParam = new object[1];
            oParam[0] = oValue;
            return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
        }

         
        public static object Invoke(object obj, string sProperty)
        {
            return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, null);
        }
    }

}
