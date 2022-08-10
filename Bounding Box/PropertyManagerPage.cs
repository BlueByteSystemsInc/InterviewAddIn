using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;

namespace Bounding_Box
{
    [ComVisibleAttribute(true)]
    public class PropertyManagerPage : PropertyManagerPage2Handler9
    {

        private SldWorks swApp1;

        PropertyManagerPage2 pmp;
        PropertyManagerPageButton pmpButton;
        PropertyManagerPageLabel pmpLabel1;
        PropertyManagerPageLabel pmpLabel2;
        PropertyManagerPageLabel pmpLabel3;

        const int buttonId = 1;
        const int label1 = 2;
        const int label2 = 3;
        const int label3 = 4;

        public PropertyManagerPage(SldWorks swApp)
        {

            string pageName = "Bounding Box Dimension";
            int longErrors = 0;
            int options = 0;

            swApp1 = swApp;

            options = (int)swPropertyManagerButtonTypes_e.swPropertyManager_OkayButton + (int)swPropertyManagerButtonTypes_e.swPropertyManager_CancelButton + (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage + (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton;

            pmp = (PropertyManagerPage2)swApp.CreatePropertyManagerPage(pageName, (int)options, this, ref longErrors);

            if (longErrors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {

                options = (int)swAddControlOptions_e.swControlOptions_Visible + (int)swAddControlOptions_e.swControlOptions_Enabled;
               pmpButton = pmp.AddControl2(buttonId, (int)swPropertyManagerPageControlType_e.swControlType_Button, "Bounding Box Dimension", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Click to get Bounding Box dimension");
                pmpLabel1 = pmp.AddControl2(label1,(int)swPropertyManagerPageControlType_e.swControlType_Label,"Width:", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Bounding Box width");
                pmpLabel2 = pmp.AddControl2(label1, (int)swPropertyManagerPageControlType_e.swControlType_Label, "Depth:", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Bounding Box Depth");
                pmpLabel3 = pmp.AddControl2(label1, (int)swPropertyManagerPageControlType_e.swControlType_Label, "Height:", (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, options, "Bounding Box Height");

            }

        }

        public void GetBoundingBoxDimension()
        {

            ModelDoc modelDoc = swApp1.ActiveDoc;

            UserUnit userUnit = modelDoc.GetUserUnit(0);

            if (modelDoc.GetType() == 2)
            {

                AssemblyDoc model = swApp1.ActiveDoc;

                object[] components;

                components = model.GetComponents(true);

                Component2 component = default(Component2);

                //////////Assembly////////////


                double[] dBox = new double[5];

                object[] bodies = null;

                Body2 body;

                double minX = 0, minY = 0, minZ = 0, maxX = 0, maxY = 0, maxZ = 0;

                bool first = true;

                double[] dCog = new double[3];

                double[] vCog = new double[3];

                object bodyInfo;

                MathUtility mathUtility = swApp1.GetMathUtility();
                MathPoint mathPoint;


                for (int i = 0; i <= components.Length - 1; i++)
                {
                    component = (Component2)components[i];
                    bodies = (object[])component.GetBodies3((int)swBodyType_e.swSolidBody, out bodyInfo);

                    for (int j = 0; j <= bodies.Length - 1; j++)
                    {
                        body = (Body2)bodies[j];

                        double x, y, z;

                        body.GetExtremePoint(1, 0, 0, out x, out y, out z);

                        dCog[0] = x; dCog[1] = y; dCog[2] = z;

                        mathPoint = mathUtility.CreatePoint(dCog);
                        mathPoint = mathPoint.MultiplyTransform(component.Transform2);

                        vCog = mathPoint.ArrayData;
                        x = vCog[0]; y = vCog[1]; z = vCog[2];

                        if (first == true || x > maxX)
                        { maxX = x; };


                        body.GetExtremePoint(-1, 0, 0, out x, out y, out z);

                        dCog[0] = x; dCog[1] = y; dCog[2] = z;

                        mathPoint = mathUtility.CreatePoint(dCog);
                        mathPoint = mathPoint.MultiplyTransform(component.Transform2);

                        vCog = mathPoint.ArrayData;
                        x = vCog[0]; y = vCog[1]; z = vCog[2];

                        if (first == true || x < minX)
                        { minX = x; };


                        body.GetExtremePoint(0, 1, 0, out x, out y, out z);

                        dCog[0] = x; dCog[1] = y; dCog[2] = z;

                        mathPoint = mathUtility.CreatePoint(dCog);
                        mathPoint = mathPoint.MultiplyTransform(component.Transform2);

                        vCog = mathPoint.ArrayData;
                        x = vCog[0]; y = vCog[1]; z = vCog[2];

                        if (first == true || y > maxY)
                        { maxY = y; };


                        body.GetExtremePoint(0, -1, 0, out x, out y, out z);

                        dCog[0] = x; dCog[1] = y; dCog[2] = z;

                        mathPoint = mathUtility.CreatePoint(dCog);
                        mathPoint = mathPoint.MultiplyTransform(component.Transform2);

                        vCog = mathPoint.ArrayData;
                        x = vCog[0]; y = vCog[1]; z = vCog[2];

                        if (first == true || y < minY)
                        { minY = y; };


                        body.GetExtremePoint(0, 0, 1, out x, out y, out z);

                        dCog[0] = x; dCog[1] = y; dCog[2] = z;

                        mathPoint = mathUtility.CreatePoint(dCog);
                        mathPoint = mathPoint.MultiplyTransform(component.Transform2);

                        vCog = mathPoint.ArrayData;
                        x = vCog[0]; y = vCog[1]; z = vCog[2];

                        if (first == true || z > maxZ)
                        { maxZ = z; };


                        body.GetExtremePoint(0, 0, -1, out x, out y, out z);

                        dCog[0] = x; dCog[1] = y; dCog[2] = z;

                        mathPoint = mathUtility.CreatePoint(dCog);
                        mathPoint = mathPoint.MultiplyTransform(component.Transform2);

                        vCog = mathPoint.ArrayData;
                        x = vCog[0]; y = vCog[1]; z = vCog[2];

                        if (first == true || z < minZ)
                        { minZ = z; };


                        first = false;

                    }

                }

                //swApp1.SendMsgToUser("Width: " + Convert.ToString(maxX - minX) + " m" + "\nDepth: " + Convert.ToString(maxZ - minZ) + " m" + "\nHeight: " + Convert.ToString(maxY - minY) + " m");

                pmpLabel1.Caption = "Width: " + Convert.ToString(userUnit.ConvertToUserUnit(maxX - minX, true, true));
                pmpLabel2.Caption = "Depth: " + Convert.ToString(userUnit.ConvertToUserUnit(maxZ - minZ, true, true));
                pmpLabel3.Caption = "Height: " + Convert.ToString(userUnit.ConvertToUserUnit(maxY - minY,true,true));

                

            }
            else if (modelDoc.GetType() == 1)
            {

                PartDoc part = swApp1.ActiveDoc;

                double[] dBox = new double[5];

                object[] bodies = null;

                Body2 body;

                double minX = 0, minY = 0, minZ = 0, maxX = 0, maxY = 0, maxZ = 0;

                bool first = true;


                bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true);


                for (int j = 0; j <= bodies.Length - 1; j++)
                {
                    body = (Body2)bodies[j];

                    double x, y, z;

                    body.GetExtremePoint(1, 0, 0, out x, out y, out z);

                    if (first == true || x > maxX)
                    { maxX = x; };


                    body.GetExtremePoint(-1, 0, 0, out x, out y, out z);


                    if (first == true || x < minX)
                    { minX = x; };


                    body.GetExtremePoint(0, 1, 0, out x, out y, out z);


                    if (first == true || y > maxY)
                    { maxY = y; };


                    body.GetExtremePoint(0, -1, 0, out x, out y, out z);


                    if (first == true || y < minY)
                    { minY = y; };


                    body.GetExtremePoint(0, 0, 1, out x, out y, out z);


                    if (first == true || z > maxZ)
                    { maxZ = z; };


                    body.GetExtremePoint(0, 0, -1, out x, out y, out z);


                    if (first == true || z < minZ)
                    { minZ = z; };


                    first = false;

                }

                //swApp1.SendMsgToUser("Width: " + Convert.ToString(maxX - minX) + " m" + "\nDepth: " + Convert.ToString(maxZ - minZ) + " m" + "\nHeight: " + Convert.ToString(maxY - minY) + " m");

                pmpLabel1.Caption = "Width: " + Convert.ToString(userUnit.ConvertToUserUnit(maxX - minX, true, true));
                pmpLabel2.Caption = "Depth: " + Convert.ToString(userUnit.ConvertToUserUnit(maxZ - minZ, true, true));
                pmpLabel3.Caption = "Height: " + Convert.ToString(userUnit.ConvertToUserUnit(maxY - minY, true, true));

            }




        }





        public void Show()
        {
            pmp.Show2(0);
        }

        #region Intefaces

        public void AfterActivation()
        {
            
        }

        public void OnClose(int Reason)
        {
            
        }

        public void AfterClose()
        {
            
        }

        public bool OnHelp()
        {
            return true;
        }

        public bool OnPreviousPage()
        {
            return true;
        }

        public bool OnNextPage()
        {
            return true;   
        }

        public bool OnPreview()
        {
            return true;
        }

        public void OnWhatsNew()
        {
            
        }

        public void OnUndo()
        {
            
        }

        public void OnRedo()
        {
            
        }

        public bool OnTabClicked(int Id)
        {
            return true;
        }

        public void OnGroupExpand(int Id, bool Expanded)
        {
            
        }

        public void OnGroupCheck(int Id, bool Checked)
        {
            
        }

        public void OnCheckboxCheck(int Id, bool Checked)
        {
            
        }

        public void OnOptionCheck(int Id)
        {
            
        }

        public void OnButtonPress(int Id)
        {
            if (Id == 1)
            { GetBoundingBoxDimension(); }
        }

        public void OnTextboxChanged(int Id, string Text)
        {
            
        }

        public void OnNumberboxChanged(int Id, double Value)
        {
            
        }

        public void OnComboboxEditChanged(int Id, string Text)
        {
            
        }

        public void OnComboboxSelectionChanged(int Id, int Item)
        {
            
        }

        public void OnListboxSelectionChanged(int Id, int Item)
        {
            
        }

        public void OnSelectionboxFocusChanged(int Id)
        {
            
        }

        public void OnSelectionboxListChanged(int Id, int Count)
        {
            
        }

        public void OnSelectionboxCalloutCreated(int Id)
        {
            
        }

        public void OnSelectionboxCalloutDestroyed(int Id)
        {
            
        }

        public bool OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            return true;
        }

        public int OnActiveXControlCreated(int Id, bool Status)
        {
            return 0;
        }

        public void OnSliderPositionChanged(int Id, double Value)
        {
            
        }

        public void OnSliderTrackingCompleted(int Id, double Value)
        {
            
        }

        public bool OnKeystroke(int Wparam, int Message, int Lparam, int Id)
        {
            return true;
        }

        public void OnPopupMenuItem(int Id)
        {
            
        }

        public void OnPopupMenuItemUpdate(int Id, ref int retval)
        {
            
        }

        public void OnGainedFocus(int Id)
        {
            
        }

        public void OnLostFocus(int Id)
        {
            
        }

        public int OnWindowFromHandleControlCreated(int Id, bool Status)
        {
            return 0;
        }

        public void OnListboxRMBUp(int Id, int PosX, int PosY)
        {
            
        }

        public void OnNumberBoxTrackingCompleted(int Id, double Value)
        {
            
        }

        #endregion
    }
}
