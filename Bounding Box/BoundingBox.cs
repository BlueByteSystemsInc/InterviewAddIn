using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using Microsoft.Win32;

namespace Bounding_Box
{
    [ComVisible(true)]
    [Guid("6929686F-1AC9-45A7-A08D-5D0436B0570B")]
    public class BoundingBox : SwAddin
    {
        public SldWorks swApp;

        private int swCookie;

        public bool ConnectToSW(object ThisSW, int Cookie)
        {

            swApp = ThisSW as SldWorks;
            swCookie = Cookie;

            bool result = swApp.SetAddinCallbackInfo(0,this,swCookie);

            int type;

            type = swApp.AddMenu((int)swDocumentTypes_e.swDocNONE, "BoundingBoxAddin", -1);
            type = swApp.AddMenu((int)swDocumentTypes_e.swDocPART, "BoundingBoxAddin", -1);
            type = swApp.AddMenu((int)swDocumentTypes_e.swDocASSEMBLY, "BoundingBoxAddin", -1);

            type = swApp.AddMenuItem4((int)swDocumentTypes_e.swDocPART, Cookie, "Show PMP Window@BoundingBoxAddin", -1, "ShowPMPWindow", "EnableMenu", "", "");
            type = swApp.AddMenuItem4((int)swDocumentTypes_e.swDocASSEMBLY, Cookie, "Show PMP Window@BoundingBoxAddin", -1, "ShowPMPWindow", "EnableMenu", "", "");

            return true;
        }

        public bool DisconnectFromSW()
        {
            return true;
        }

        public void ShowPMPWindow()
        {
            
            PropertyManagerPage pmp = new PropertyManagerPage(swApp);
            pmp.Show();
        
        }


        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {


            string path = String.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            RegistryKey key = Registry.LocalMachine.CreateSubKey(path);

            key.SetValue(null, 1);
            key.SetValue("Title", "Bounding Box");
            key.SetValue("Description", "To get the Bounding Box dimension");




        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {

            string path = String.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            Registry.LocalMachine.DeleteSubKey(path);


        }

        #endregion


    }
}
