
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using Microsoft.Office.Tools;
using System.Reflection;
using System.Deployment.Application;

namespace XLAppAddIn {
    [ComVisible(true)]
    public interface IAddInUtilities
    {

        void SetCustomPanePositionWhenFloating(int x, int y, string ctpName = "", Microsoft.Office.Tools.CustomTaskPane customTaskPane = null);

        void ImportData();
        void ShowMessageBox();
        void HideUserControl();
        void ShowUserControl();
        void ShowOrHideUserControl();
        void ShowRibbonAddinTab();
        int GetUserControlWidth();
        bool GetUserControlIsVisible();
        void AdjustComboBoxLine(string indText);
        void ShowAppVertBar();
        void ToggleAppVerifProjet();
        void ToggleCopyPasteRibbon(string enabled);

        string GetClickOnceLocation();
        string CurrentVersion();
        void InstallUpdateSyncWithInfo();
        //bool GetIsAddIn(); est une static bool ici, bas, donc pourrait aller dans
        // une autre classe à part car ne peut être callé dans Excel
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]   

    public class AddInUtilities : IAddInUtilities
    {

        public void InstallUpdateSyncWithInfo() {
            // https://msdn.microsoft.com/en-us/library/ms404263.aspx
            UpdateCheckInfo info = null;

            if (ApplicationDeployment.IsNetworkDeployed) {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                try {
                    info = ad.CheckForDetailedUpdate();

                } catch (DeploymentDownloadException dde) {
                    MessageBox.Show("The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error: " + dde.Message);
                    return;
                } catch (InvalidDeploymentException ide) {
                    MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message);
                    return;
                } catch (InvalidOperationException ioe) {
                    MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message);
                    return;
                }

                if (info.UpdateAvailable) {
                    Boolean doUpdate = true;

                    if (!info.IsUpdateRequired) {
                        DialogResult dr = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available", MessageBoxButtons.OKCancel);
                        if (!(DialogResult.OK == dr)) {
                            doUpdate = false;
                        }
                    } else {
                        // Display a message that the app MUST reboot. Display the minimum required version.
                        MessageBox.Show("This application has detected a mandatory update from your current " +
                            "version to version " + info.MinimumRequiredVersion.ToString() +
                            ". The application will now install the update and restart.",
                            "Update Available", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (doUpdate) {
                        try {
                            ad.Update();
                            MessageBox.Show("The application has been upgraded, and will now restart.");
                            Application.Restart();
                        } catch (DeploymentDownloadException dde) {
                            MessageBox.Show("Cannot install the latest version of the application. \n\nPlease check your network connection, or try again later. Error: " + dde);
                            return;
                        }
                    }
                }
            }
        }


        public string CurrentVersion() {
            // How to get current the product version in C#?
            // Just give the reference to System.Deployment.Application and though it wont work in developement of the visual studio but it will work once the application is deployed.

            ////using System.Deployment.Application;
            ////using System.Reflection; 
            return ApplicationDeployment.IsNetworkDeployed
                       ? ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString() // retourne la bonne version en exécution
                       : Assembly.GetExecutingAssembly().GetName().Version.ToString(); //le 2e retourne : 1.0.0.0
            // si la version ici ne match pas celui indiqué dans le raccourci du desktop (et startmenu), on  copiera les dossiers -Projets, -Projets BackUp, ainsi que le fichier -Importation Bordereau du dossier Resources et le fichier -logo.png du dossier Images 
        }

        public string GetClickOnceLocation() {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;
            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            return Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
        }


        // This method tries to write a string to cell A1 in the active worksheet.
        public void ImportData()
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            if (activeWorksheet != null)
            {
                Excel.Range range1 = activeWorksheet.get_Range("A1", System.Type.Missing);
                range1.Value2 = "This is my data";
            }
        }

        public void ShowMessageBox()
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (activeWorksheet != null)
            {
                System.Windows.Forms.MessageBox.Show("I am called from C# COM add-in");
            }

        }
        public void ToggleCopyPasteRibbon(string enabled = "False")
        {
            ManageTaskPaneRibbon.ToggleCopyPasteRibbon(enabled);
        }

        public void ShowRibbonAddinTab()
        {
            ManageTaskPaneRibbon.ShowRibbonAddinTab();
        }

        public void HideUserControl()
        {
            ManageTaskPaneRibbon.HideUserControl();
        }
        public void ShowUserControl()
        {
            ManageTaskPaneRibbon.ShowUserControl();
        }

        public void SetCustomPanePositionWhenFloating(int x, int y, string ctpName = "", Microsoft.Office.Tools.CustomTaskPane customTaskPane = null) {
            //var oldDockPosition = customTaskPane.DockPosition;
            //var oldVisibleState = customTaskPane.Visible;

            //////Globals.ThisAddIn.myUserControlWPF.RemoveResizeEvent();
            //Application.DoEvents();

            if (customTaskPane == null && ctpName.Length == 0)
                customTaskPane = Globals.ThisAddIn.myCustomTaskPaneWPF;
            else if (ctpName.Length > 0)
                customTaskPane = getCTPByName(ctpName);

            customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;



            customTaskPane.Visible = true; //The task pane must be visible to set its position
            
            IntPtr window = NativeMethods.FindWindowW("MsoCommandBar", customTaskPane.Title); //MLHIDE
            if (window == null) return;

            if (!MoveWindow(window, x, y, customTaskPane.Width, customTaskPane.Height, true)) {
                //throw new Exception();
            }

            //////Globals.ThisAddIn.myUserControlWPF.EnableResizeEvent();
            //Application.DoEvents();

            //customTaskPane.Visible = oldVisibleState;
            //customTaskPane.DockPosition = oldDockPosition;
        }

        private CustomTaskPane getCTPByName(string ctpName) {
            switch (ctpName) {
                case "XLApp - Estimation imposition":

                    return Globals.ThisAddIn.myCustomTaskPaneWPF;
                case "test2":

                    return Globals.ThisAddIn.myCustomTaskPaneWPF;
                default:
                    return Globals.ThisAddIn.myCustomTaskPaneWPF;
            }
        }

        //[DllImport("user32.dll", EntryPoint = "FindWindowW")]
        //public static extern System.IntPtr FindWindowW([System.Runtime.InteropServices.InAttribute()] [System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string lpClassName, [System.Runtime.InteropServices.InAttribute()] [System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "MoveWindow")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool MoveWindow([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, int X, int Y, int nWidth, int nHeight, [System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)] bool bRepaint);

        internal static void InitiateFirstLaunch()
        {

            // Validation1 : vérifier si les raccourcis sont sur le bureau :
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (File.Exists(userProfile + "\\Desktop\\" + "XLApp" + ".lnk")) return;

            DialogResult result1 = MessageBox.Show("Démarrer le lancement d'initialisation de l'application?",
                                                    "XLApp",
                                                    MessageBoxButtons.YesNo,
                                                    MessageBoxIcon.Question,
                                                    MessageBoxDefaultButton.Button1);

            if (result1 == DialogResult.Yes)
            {
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;

                // Avant le démarrement du lancement d'initialisation de l'application, 
                // on pourrait activer la clé de registre qui coche et permet donc l'accès au trust model VBA
                // https://blogs.msdn.microsoft.com/cristib/2012/02/29/vba-how-to-programmatically-enable-access-to-the-vba-object-model-using-macros/

                // Cela peut se faire en ajoutant les composantes VB au document excel actif et en runnant la macro qui quittera la scéance
                // active d'Excel, ou en codant directement ds c#, mais la scéance doit être fermer.






                Excel.Application xlApp;

                Excel.Workbook xlWorkBook;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                // http://stackoverflow.com/questions/1600502/programmatically-configuring-ms-words-trust-center-settings-using-c-sharp
                // a very simple way of opening an .xls file containing macros, without messing with the registry or Excel's trust settings. :
                xlApp.FileValidation = Microsoft.Office.Core.MsoFileValidationMode.msoFileValidationSkip;
                // fin a very simple way :)

                //Get the assembly information
                System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
                //Location is where the assembly is run from 
                string assemblyLocation = assemblyInfo.Location;
                //CodeBase is the location of the ClickOnce deployment files
                Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
                string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
                //xlApp.Caption = "";
                //                                                                   readonly=true
                xlWorkBook = xlApp.Workbooks.Open(ClickOnceLocation + "\\XLApp.dll", 0, true, 5, "VelvetSweatshop911", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", false, false, 0, false, 1, 0);
                xlApp.Caption = "";
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                //xlApp.Visible = true;
                // xlApp = null;

                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
            }
        }

        internal static void LaunchApp()
        {
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            // http://stackoverflow.com/questions/1600502/programmatically-configuring-ms-words-trust-center-settings-using-c-sharp
            //"a very simple way of opening an .xls file containing macros, without messing with the registry or Excel's trust settings." :
            xlApp.FileValidation = Microsoft.Office.Core.MsoFileValidationMode.msoFileValidationSkip;
            // fin a very simple way :)


            xlApp.EnableCancelKey = Excel.XlEnableCancelKey.xlDisabled;
            xlApp.CalculationInterruptKey = Excel.XlCalculationInterruptKey.xlNoKey;

            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;
            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
          
            //                                                                   readonly=true //pw
            xlWorkBook = xlApp.Workbooks.Open(ClickOnceLocation + "\\XLApp.dll", 0, true, 5, "False", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", false, false, 0, false, 1, 0);
            
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //xlApp.Visible = true;
            // xlApp = null;

            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
        }

        public void ShowOrHideUserControl()
        {
            ManageTaskPaneRibbon.ShowOrHideUserControl();
        }
        public int GetUserControlWidth()
        {
            return ManageTaskPaneRibbon.GetUserControlWidth();
        }
        public bool GetUserControlIsVisible()
        {
            return Globals.ThisAddIn.TaskPane.Visible;
        }
        public void AdjustComboBoxLine(string indText)
        {
            ManageTaskPaneRibbon.AdjustComboBoxLine(indText);
        }
        public void ShowAppVertBar() //testing 11-12-2016 from vba
        {
            Globals.ThisAddIn.TaskPane.Visible = !(Globals.ThisAddIn.TaskPane.Visible);
            //Globals.ThisAddIn.Application.Run("resizeWindow");
        }
        public void ToggleAppVerifProjet()
        {
            Globals.ThisAddIn.TaskPaneVerifProjet.Visible = !(Globals.ThisAddIn.TaskPaneVerifProjet.Visible);
        }
        //Methods callés dand c# existants ds VBA :
        public static bool GetIsAddIn(out bool isMyApp)
        {
             try
             {
                isMyApp = (Globals.ThisAddIn.Application.Caption.IndexOf("XLCie") > -1);
                return (Globals.ThisAddIn.Application.ActiveWorkbook.IsAddin);
            }
             catch
             {
                // Si erreur, c'est parce que c'est un addin (par essai-erreur)
                isMyApp = (Globals.ThisAddIn.Application.Caption.IndexOf("XLCie") > -1);
                return true; 
             }
        }
        public static void UnConnectAddin()
        {
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office\\Excel\\Addins\\XLAppAddIn", true);
            if (registryKey != null)
            {
                registryKey.SetValue("LoadBehavior", 2); //unload --- https://msdn.microsoft.com/en-us/library/bb386106.aspx
            }
        }

        //GROUPE Saisie des items
        public static void insertArticle()
        {
            Globals.ThisAddIn.Application.Run("test1");
        }
        public static void insertArticleAss()
        {
            Globals.ThisAddIn.Application.Run("InsertionMultiple");
        }
        public static void insertAss()
        {
            Globals.ThisAddIn.Application.Run("Montrerass");
        }
        public static void OrderArt()
        {
            Globals.ThisAddIn.Application.Run("OrdonnerArticles");
        }
        public static void ShowFunctions()
        {
            Globals.ThisAddIn.Application.Run("ShowFunctions");
        }
        public static void TextChangedComboBoxLignes(string cSharpIndex)
        {
            Globals.ThisAddIn.Application.Run("SynconizeLineComboBoxs", cSharpIndex);
        }
        public static void InsertLineSaisieBanque()
        {
            //Globals.ThisAddIn.Application.Run("InsertLineSaisieBanque");
            //Application.DoEvents();
            SendKeys.Send("{INSERT}");
        }
        public static void SupprimerLigne()
        {
            Globals.ThisAddIn.Application.Run("SupprimerLigne");
        }
        public static void SupprimerProduits()
        {
            Globals.ThisAddIn.Application.Run("SupprimerProduits");
        }
        public static void SupprimerLignesMultiples()  // supprimer 1 Article
        {
            Globals.ThisAddIn.Application.Run("SupprimerLignesMultiples");
        }
        public static void Macroretablirarticles()
        {
            Globals.ThisAddIn.Application.Run("Macrorétablirarticles");
        }
        public static void ShowBordImp()
        {
            Globals.ThisAddIn.Application.Run("ShowBordImp");
        }

        // fin groupe Saisie

        //NavigatiON:   '1:first 2:previos 3:next 4:last
        public static void GotoPremier()
        {
            Globals.ThisAddIn.Application.Run("GoToArticleFromCSharp", 1);   //to create in vba
        }
        public static void GotoPrec()
        {
            Globals.ThisAddIn.Application.Run("GoToArticleFromCSharp", 2);         //to create in vba
        }
        public static void GotoSuiv()
        {
            Globals.ThisAddIn.Application.Run("GoToArticleFromCSharp", 3);      //to create in vba
        }
        public static void GotoLast()
        {
            Globals.ThisAddIn.Application.Run("GoToArticleFromCSharp", 4);   //to create in vba
        }
        public static void MODART()
        {
            Globals.ThisAddIn.Application.Run("MODART");
        }
        public static void GotoArticle()
        {
            Globals.ThisAddIn.Application.Run("GotoArticle");
        }
        // fin navigation


        //gestionnaire proj
        public static void NouveauProjet()
        {
            Globals.ThisAddIn.Application.Run("NouveauProjet");
        }
        public static void OuvrirProjet()
        {
            Globals.ThisAddIn.Application.Run("OuvrirProj");
        }
        public static void EnregProj()
        {
            Globals.ThisAddIn.Application.Run("Create_Tables_XLApp");
        }
        public static void EnregProjSous()
        {
            Globals.ThisAddIn.Application.Run("Create_Tables_XLApp_As_New");
        }
        public static void FermerProjet()
        {
            Globals.ThisAddIn.Application.Run("FermerProjet");
        }

        internal static void IMPRBORD()
        {
            Globals.ThisAddIn.Application.Run("IMPRBORD");
        }

        //fin gestionnaire
        //déplacement
        public static void Coller()
        {
            Globals.ThisAddIn.Application.Run("DoPaste");
        }

        internal static void SignatureAuto()
        {
            Globals.ThisAddIn.Application.Run("SignatureAuto");
        }

        internal static void ZoneImpAuto()
        {
            Globals.ThisAddIn.Application.Run("ZoneImpAuto");
        }

        internal static void ListerRessources()
        {
            Globals.ThisAddIn.Application.Run("QueryOnOpenWBprRessources");
        }

        internal static void ClickMEP()
        {
            Globals.ThisAddIn.Application.Run("ZoneImpAuto");
        }

        internal static void AMCoutU()
        {
            Globals.ThisAddIn.Application.Run("AMCoutU");
        }

        public static void Copier()
        {
            Globals.ThisAddIn.Application.Run("docopy");
        }
        public static void Couper()
        {
            Globals.ThisAddIn.Application.Run("DoCut");
        }
        public static void ClearClip()
        {
            Globals.ThisAddIn.Application.Run("ClearClipBoard");
        }
        //fin déplacement
        public static void FilterProduitsByLetterAndGoToFirstLetter(string letter)
        {
            Globals.ThisAddIn.Application.Run("FilterProduitsByLetterAndGoToFirstLetter", letter);
        }
    }

    }
