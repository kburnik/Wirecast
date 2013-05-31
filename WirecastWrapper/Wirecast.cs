/*
 Wirecast 4.3 .NET Wrapper v 0.1 by Kristijan Burnik <kristijanburnik@gmail.com>
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace WirecastWrapper
{

  

    public enum CompareMethod
    {
        ExactMatch = 0,
        Contains = 1,
        CaseInsensitive = 2,
        CaseInsensitiveContains = 3
    }

    public enum TransitionSpeed
    {
        slowest, slow, normal, fast, fastest
    }

    public enum ActiveTransitionIndex
    {
        Null = 0,
        FirstPopup = 1,
        SecondPopup = 2,
        ThirdPopup = 3
    }


    public class Wirecast
    {
        public static readonly int TOTAL_LAYER_COUNT = 5;
        public static readonly int DEFAULT_LAYER_INDEX = 2; // third layer is default

        private static Wirecast _instance = new Wirecast();
        public static Wirecast Instance
        {
            get
            {
                return _instance;
            }
        }

        

        //////////////////////////////////////////////////////////////////

        private Wirecast()
        {

        }

        private object oWirecast { get { return GetWirecast() ; } }

        [MTAThread]
        public WirecastDocument DocumentByIndex(int index)
        {
            return WirecastDocument.CreateFromObject(index, Late.Invoke(oWirecast, "DocumentByIndex", index));
        }

        public WirecastDocument DocumentByName(string name, CompareMethod compareMethod)
        {
            return new WirecastDocument(Late.Invoke(oWirecast, "DocumentByName",
                new object[] { name, (int)compareMethod }));
        }



        /// <summary>
        /// Get the active Wirecast object.
        /// </summary>
        /// <returns>the active wirecast object</returns>
        /// 
        [MTAThread]
        private static object GetWirecast()
        {

            try
            {
                return Marshal.GetActiveObject("Wirecast.Application");
            }
            catch
            {
                Type objClassType = Type.GetTypeFromProgID("Wirecast.Application");
                return Activator.CreateInstance(objClassType);
            }
        }


        public static void StartWirecast() {
            GetWirecast();
        }

    }
}
