using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Forms;

//using System.IO;


namespace XLAppAddIn
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]

    public class ConvertImage : System.Windows.Forms.AxHost //pour commandBar image
    {
        private ConvertImage()
            : base(null)
        {
        }

        public static stdole.IPictureDisp Convert
            (System.Drawing.Image image)
        {
            return (stdole.IPictureDisp)System.
                Windows.Forms.AxHost
                .GetIPictureDispFromPicture(image);
        }

    }

    public partial class NativeMethods {

        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "FindWindowW")]

        public static extern System.IntPtr FindWindowW([System.Runtime.InteropServices.InAttribute()] [System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string lpClassName, [System.Runtime.InteropServices.InAttribute()] [System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string lpWindowName);

    }
    public partial class ThisAddIn
    {

        private Office.CommandBarButton buttonOne; //pour custom image commandBar

        private InterfaceVert myUserControl1;
        private InteractSousTrait myUserControlInterfaceSousTrait;
        private VerifProjet myUserControlVerificationProjet;
        private VerifProjet myUserControlInterfaceData;

        public UserControl2 myUserControlWPF; // TESTWPF https://msdn.microsoft.com/en-ca/library/bb772076.aspx https://msdn.microsoft.com/en-ca/library/bb384311.aspx //UserControl2.cs //UserControl1.xaml
                                              //// ajouter WPF Usercontrol type WPF, faire du drag and drop avec les outils, générer projet, ajouter usercontrol windows forms, mettre le code pour le taskpane, drag and drop de usercontrol wpf à usercontrol windowsforms

        private int wpfPaneWidth = 780;
        private int wpfPaneHeight = 325;


        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane; //was private, test pour userControl Resize - SQL
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPaneInterfaceSousTrait;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPaneVerificationProjet;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPaneInterfaceData;

        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPaneWPFEstImp; // TESTWPF

        private AddInUtilities utilities;

        // taskpane #1 : InterfaceVert
        public Microsoft.Office.Tools.CustomTaskPane TaskPaneInterfaceVert {
            get {
                return myCustomTaskPane;
            }
        }
        //test taskpane #2 : InteractSousTrait
        public Microsoft.Office.Tools.CustomTaskPane TaskPaneInterfaceSousTrait {
            get {
                return myCustomTaskPaneInterfaceSousTrait;
            }
        }
        //test taskpane #3 : VerifProjet
        public Microsoft.Office.Tools.CustomTaskPane TaskPaneVerifProjet {
            get {
                return myCustomTaskPaneVerificationProjet;
            }
        }
        //test taskpane #3 : VerifProjetTemp
        public Microsoft.Office.Tools.CustomTaskPane TaskPaneInterfaceData {
            get {
                return myCustomTaskPaneInterfaceData;
            }
        }
        //test taskpane #4 : VerifProjetTemp
        public Microsoft.Office.Tools.CustomTaskPane TaskPaneEstImposWPF {
            get {
                return myCustomTaskPaneWPFEstImp;
            }
        }


        protected override object RequestComAddInAutomationService()
        {

            {

                try
                {
                    if (utilities == null)
                        utilities = new AddInUtilities();
                }
                catch
                {
                }
                return utilities;

            }

        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {


            try
            {
                if (this.Application.Caption.IndexOf("XLCie") > -1)
                {
                    var width = 0;



                    //USERCONTROL DE GAUCHE AVEC TABLE LAYOUT PANNEL
                    myUserControl1 = new InterfaceVert();
                    //set Width du au UserControl Docking

                    width = myUserControl1.Width;
                    myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "XLApp");

                    //set Width du au UserControl Docking
                    myCustomTaskPane.Width = width;

                    myCustomTaskPane.DockPosition =
                         Office.MsoCTPDockPosition.msoCTPDockPositionLeft;

                    myCustomTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

                    myCustomTaskPane.VisibleChanged +=
                        new EventHandler(myCustomTaskPane_VisibleChanged);

                    //////myCustomTaskPane.Visible = true;
                    Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked = myCustomTaskPane.Visible;
                    if (this.Application.Caption.IndexOf("XLCie") > -1) this.Application.Run("ShowOrHideUserControlCheckBoxInCSharp"); //met myUserControl1 visible
                                                                                                                                       //FIN USERCONTROL DE GAUCHE AVEC TABLE LAYOUT PANNEL

                    //TEST NEW USERCONTROL InterfaceData :
                    myUserControlInterfaceData = new VerifProjet();
                    width = myUserControlInterfaceData.Width;
                    myCustomTaskPaneInterfaceData = this.CustomTaskPanes.Add(myUserControlInterfaceData, "Data applicatif");

                    //set Width du au UserControl Docking
                    myCustomTaskPaneInterfaceData.Width = width;

                    myCustomTaskPaneInterfaceData.DockPosition =
                         Office.MsoCTPDockPosition.msoCTPDockPositionLeft;

                    myCustomTaskPaneInterfaceData.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

                    myCustomTaskPaneInterfaceData.VisibleChanged +=
                        new EventHandler(myCustomTaskPaneInterfaceData_VisibleChanged);
                    //myCustomTaskPaneVerificationProjetTemp.VisibleChanged +=
                    //    new EventHandler(myCustomTaskPaneVerificationProjetTemp_VisibleChanged);
                    //FIN TEST


                    //TEST InteractSousTraitant :
                    myUserControlInterfaceSousTrait = new InteractSousTrait();
                    width = myUserControlInterfaceSousTrait.Width;
                    myCustomTaskPaneInterfaceSousTrait = this.CustomTaskPanes.Add(myUserControlInterfaceSousTrait, " ");

                    //set Width du au UserControl Docking
                    myCustomTaskPaneInterfaceSousTrait.Width = width;

                    myCustomTaskPaneInterfaceSousTrait.DockPosition =
                         Office.MsoCTPDockPosition.msoCTPDockPositionLeft;

                    myCustomTaskPaneInterfaceSousTrait.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

                    myCustomTaskPaneInterfaceSousTrait.VisibleChanged +=
                        new EventHandler(myCustomTaskPaneInterfaceSousTrait_VisibleChanged);
                    //FIN TEST


                    //TEST NEW USERCONTROL VERIFPROJET :
                    myUserControlVerificationProjet = new VerifProjet();
                    width = myUserControlVerificationProjet.Width;
                    myCustomTaskPaneVerificationProjet = this.CustomTaskPanes.Add(myUserControlVerificationProjet, "Vérification rapide");

                    //set Width du au UserControl Docking
                    myCustomTaskPaneVerificationProjet.Width = width;

                    myCustomTaskPaneVerificationProjet.DockPosition =
                         Office.MsoCTPDockPosition.msoCTPDockPositionRight;

                    myCustomTaskPaneVerificationProjet.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

                    myCustomTaskPaneVerificationProjet.VisibleChanged +=
                        new EventHandler(myCustomTaskPaneVerificationProjet_VisibleChanged);
                    //FIN TEST


                    //// TEST USERCONTROL WPF
                    myUserControlWPF = new UserControl2(); // UserContron2.cs
                    width = myUserControlWPF.Width;
                    int height = myUserControlWPF.Height;
                    myCustomTaskPaneWPFEstImp = this.CustomTaskPanes.Add(myUserControlWPF, "XLApp - Estimation imposition");

                    myCustomTaskPaneWPFEstImp.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
                    //myUserControlWPF.SizeChanged -= UserControl2_SizeChanged;
                    myCustomTaskPaneWPFEstImp.Height = height + 45;
                    myCustomTaskPaneWPFEstImp.Width = width + 15;

                    myCustomTaskPaneWPFEstImp.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

                    myCustomTaskPaneWPFEstImp.Control.SizeChanged += new EventHandler(Control_SizeChanged);
                    //myCustomTaskPaneWPF.Control.MinimumSize = new System.Drawing.Size(100, 100);
                    //myCustomTaskPaneWPF.Control.MaximumSize = new System.Drawing.Size(myCustomTaskPaneWPF.Width, myCustomTaskPaneWPF.Height);


                    //SetCustomPanePositionWhenFloating(myCustomTaskPaneWPF, (int)this.Application.Left, (int)this.Application.Top);
                    //myCustomTaskPaneWPF.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

                    //myCustomTaskPaneWPF.VisibleChanged +=
                    //        new EventHandler(myCustomTaskPaneWPF_VisibleChanged);
                    //FIN TEST

                    //pour image custom:
                   defineCustomCommandBarImages();
                }
                else
                {
                    AddInUtilities.UnConnectAddin();
                    //AddInUtilities.InitiateFirstLaunch();
                }
            }
            catch
            {
            }

        }

        private void Control_SizeChanged(object sender, EventArgs e) {

            if (myCustomTaskPaneWPFEstImp.Height > wpfPaneHeight && myCustomTaskPaneWPFEstImp.Width > wpfPaneWidth) {
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;
                Globals.ThisAddIn.Application.SendKeys("{ESC}", true);
                //myCustomTaskPaneWPF.Control.SizeChanged -= new EventHandler(Control_SizeChanged);
                myCustomTaskPaneWPFEstImp.Height = wpfPaneHeight;
                myCustomTaskPaneWPFEstImp.Width = wpfPaneWidth;
                //System.Windows.Forms.Application.DoEvents();
                //DoMouseUp();
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;

                //myCustomTaskPaneWPF.Control.SizeChanged += new EventHandler(Control_SizeChanged);
            } else if (myCustomTaskPaneWPFEstImp.Height > wpfPaneHeight) {
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;
                Globals.ThisAddIn.Application.SendKeys("{ESC}", true);
                //myCustomTaskPaneWPF.Control.SizeChanged -= new EventHandler(Control_SizeChanged);
                myCustomTaskPaneWPFEstImp.Height = wpfPaneHeight;
                //System.Windows.Forms.Application.DoEvents();
                //DoMouseUp();
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;

                //myCustomTaskPaneWPF.Control.SizeChanged += new EventHandler(Control_SizeChanged);
            } else if (myCustomTaskPaneWPFEstImp.Width > wpfPaneWidth) {
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;
                Globals.ThisAddIn.Application.SendKeys("{ESC}", true);

                //myCustomTaskPaneWPF.Control.SizeChanged -= new EventHandler(Control_SizeChanged);
                myCustomTaskPaneWPFEstImp.Width = wpfPaneWidth;
                //System.Windows.Forms.Application.DoEvents();
                //DoMouseUp();
                //myCustomTaskPaneWPF.Control.SizeChanged += new EventHandler(Control_SizeChanged);
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
            }



        }
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        public void DoMouseUp() {
            //Call the imported function with the cursor's current position
            int X = Cursor.Position.X;
            int Y = Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTUP | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }

        //private void SetCustomPanePositionWhenFloating(Microsoft.Office.Tools.CustomTaskPane customTaskPane, int x, int y) {
        //    //var oldDockPosition = customTaskPane.DockPosition;
        //    //var oldVisibleState = customTaskPane.Visible;

        //    customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
        //    customTaskPane.Visible = true; //The task pane must be visible to set its position

        //    IntPtr window = NativeMethods.FindWindowW("MsoCommandBar", customTaskPane.Title); //MLHIDE
        //    if (window == null) return;

        //    if (!MoveWindow(window, x, y, customTaskPane.Width, customTaskPane.Height, true)) {
        //        throw new Exception();
        //    }
        //    //customTaskPane.Visible = oldVisibleState;
        //    //customTaskPane.DockPosition = oldDockPosition;
        //}



        private void SetCustomPaneSizeWhenFloating(Microsoft.Office.Tools.CustomTaskPane customTaskPane, int width, int height) {
            var oldDockPosition = customTaskPane.DockPosition;

            customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
            customTaskPane.Width = width;
            customTaskPane.Height = height;

            customTaskPane.DockPosition = oldDockPosition;
        }
        private void defineCustomCommandBarImages()
        {
            if (this.Application.Caption.IndexOf("XLCie") == -1) return;
            // Application.CommandBars(1).FindControl(msoControlButton, Tag:= 2172)
            // Application.CommandBars("Cell").FindControl(msoControlButton, Tag:= 2172)
            // Application.CommandBars("Row").FindControl(msoControlButton, Tag:= 2172)
            // Application.CommandBars("Desktop").FindControl(msoControlButton, Tag:= 2172)

            buttonOne = (Office.CommandBarButton)Application.CommandBars["Cell"].FindControl(
                    Office.MsoControlType.msoControlButton, missing, 2172, false, null);

            buttonOne.Picture = getImage(); /// *****

            buttonOne = (Office.CommandBarButton)Application.CommandBars["Row"].FindControl(
                Office.MsoControlType.msoControlButton, missing, 2172, false, null);

            buttonOne.Picture = getImage(); /// *****

            buttonOne = (Office.CommandBarButton)Application.CommandBars["Desktop"].FindControl(
    Office.MsoControlType.msoControlButton, missing, 2172, false, null);

            buttonOne.Picture = getImage(); /// *****

            buttonOne = (Office.CommandBarButton)Application.CommandBars[1].FindControl(
Office.MsoControlType.msoControlButton, missing, 2172, false, null);

            buttonOne.Picture = getImage(); /// *****


        }
        private stdole.IPictureDisp getImage()
        {
            stdole.IPictureDisp tempImage = null;
            try
            {
                System.Drawing.Image newIcon =
                    Properties.Resources.CC16161;    // changé .Icon par .Image avec .png - marche bien avec png 1616

                System.Windows.Forms.ImageList newImageList =
                    new System.Windows.Forms.ImageList();
                newImageList.Images.Add(newIcon);
                tempImage = ConvertImage.Convert(newImageList.Images[0]);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return tempImage;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office\\Excel\\Addins\\XLAppAddIn", true);
            if (registryKey != null)
            {
                registryKey.SetValue("LoadBehavior", 2); //unload --- https://msdn.microsoft.com/en-us/library/bb386106.aspx
            }
        }

        private void myCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked = myCustomTaskPane.Visible;
            try
            {
                Globals.ThisAddIn.Application.Run("ShowOrHideUserControlCheckBoxInCSharp", myCustomTaskPane.Visible); //passé paramètre ici
                Globals.ThisAddIn.Application.Run("resizeWindow");
            }
            catch
            {

            }


        }

        //myCustomTaskPaneInterfaceSousTrait_VisibleChanged
        private void myCustomTaskPaneInterfaceSousTrait_VisibleChanged(object sender, EventArgs e)
        {
            if (!myCustomTaskPaneInterfaceSousTrait.Visible)
            {

                myCustomTaskPaneInterfaceData.VisibleChanged -= myCustomTaskPaneInterfaceData_VisibleChanged;
                myCustomTaskPaneInterfaceData.Visible = false;

                myUserControl1.EnableButtonInterfaceVert("buttonSaisie", "tableLayoutPanel1"); //PerformClick sur buttonSaisie

                myCustomTaskPaneInterfaceData.VisibleChanged += myCustomTaskPaneInterfaceData_VisibleChanged;

            }

        }
        private void myCustomTaskPaneInterfaceData_VisibleChanged(object sender, EventArgs e)
        {
            if (!myCustomTaskPaneInterfaceData.Visible)
            {
                //InterfaceVert myForm = new InterfaceVert();
                //myForm.buttonEst.Enabled = true;
                //myForm.buttonSaisie.Enabled = true;

                myCustomTaskPaneInterfaceSousTrait.VisibleChanged -= myCustomTaskPaneInterfaceSousTrait_VisibleChanged;
                myCustomTaskPaneInterfaceSousTrait.Visible = false;

                myUserControl1.EnableButtonInterfaceVert("buttonSaisie", "tableLayoutPanel1");  //PerformClick sur buttonSaisie

                myCustomTaskPaneInterfaceSousTrait.VisibleChanged += myCustomTaskPaneInterfaceSousTrait_VisibleChanged;

            }
        }

        //myCustomTaskPaneVerificationProjet_VisibleChanged
        private void myCustomTaskPaneVerificationProjet_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.toggleButtonVerif.Checked = myCustomTaskPaneVerificationProjet.Visible;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}