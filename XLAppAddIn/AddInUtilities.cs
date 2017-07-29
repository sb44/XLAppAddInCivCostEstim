
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using Microsoft.Office.Tools;
using System.Reflection;
using System.Deployment.Application;
using System.Security;
using System.Security.Policy;
using System.Security.Permissions;
using System.Diagnostics;
using System.Collections.Generic;

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
        string InstallUpdateSyncWithInfo();

        //bool GetIsAddIn(); est une static bool ici, bas, donc pourrait aller dans
        // une autre classe à part car ne peut être callé dans Excel
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]   

    public class AddInUtilities : IAddInUtilities
    {
   //     https://blogs.msdn.microsoft.com/krimakey/2008/04/18/click-once-forced-updates-in-vsto-ii-a-fuller-solution/
        public string InstallUpdateSyncWithInfo() {
            // https://msdn.microsoft.com/en-us/library/ms404263.aspx
            UpdateCheckInfo info = null;

            if (!ApplicationDeployment.IsNetworkDeployed)
                return "La version actuelle n'est pas déployé en réseau ou est une version de développement.";

                if (ApplicationDeployment.IsNetworkDeployed) {

                Assembly addinAssembly = Assembly.GetExecutingAssembly();

                string CachePath = addinAssembly.CodeBase.Substring(0, addinAssembly.CodeBase.Length -
                    System.IO.Path.GetFileName(addinAssembly.CodeBase).Length);

                ApplicationDeployment CurrentDep = ApplicationDeployment.CurrentDeployment;

                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                // https://blogs.msdn.microsoft.com/krimakey/2008/04/10/click-once-forced-updates-in-vsto-some-things-we-dont-recommend-using-that-you-might-consider-anyway/
                ApplicationIdentity appId = new ApplicationIdentity(ad.UpdatedApplicationFullName);

                PermissionSet unrestrictedPerms = new PermissionSet(PermissionState.Unrestricted);

                ApplicationTrust appTrust = new ApplicationTrust(appId) {
                    DefaultGrantSet = new PolicyStatement(unrestrictedPerms),
                    IsApplicationTrustedToRun = true,
                    Persist = true
                };

                ApplicationSecurityManager.UserApplicationTrusts.Add(appTrust);

                try {
                    info = ad.CheckForDetailedUpdate();

                } catch (DeploymentDownloadException dde) {
                    //MessageBox.Show("The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error: " + dde.Message);
                    //MessageBox.Show("La nouvelle version de l'application ne peut être télécharger en ce moment. \n\nVeuillez vérifier votre connection, ou réessayer plus tard. Erreur: " + dde.Message);
                    return "La nouvelle version de l'application ne peut être télécharger en ce moment. \n\nVeuillez vérifier votre connection, ou réessayer plus tard. Erreur: " + dde.Message;
                } catch (InvalidDeploymentException ide) {
                    //MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message);
                    //MessageBox.Show("Impossible de vérifier pour une nouvelle version de l'application. Le déploiement ClickOnce de l'application est corrompue. Veuillez redéployez l'application et réessayer. Erreur: " + ide.Message);
                    return "Impossible de vérifier pour une nouvelle version de l'application. Le déploiement ClickOnce de l'application est corrompue. Veuillez redéployez l'application et réessayer. Erreur: " + ide.Message;
                } catch (InvalidOperationException ioe) {
                    //MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message);
                    //MessageBox.Show("Cet application ne peut être mise à jour. Ce n'est vraisemblablement pas une application ClickOnce. Erreur: " + ioe.Message);
                    return "Cet application ne peut être mise à jour. Ce n'est vraisemblablement pas une application ClickOnce. Erreur: " + ioe.Message;
                }

                if (!info.UpdateAvailable)
                    return "La version actuelle (" + (DateTime.Now.Year % 100).ToString() + "." + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString() + ") est à jour.";

               if (info.UpdateAvailable) {
                    Boolean doUpdate = true;

                  //  string test = CurrentDep.UpdatedVersion.ToString();

                    if (!info.IsUpdateRequired) {
                        //DialogResult dr = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available", MessageBoxButtons.OKCancel);
                        DialogResult dr = MessageBox.Show("Une mise à jour de l'application est disponible. Souhaitez-vous l'exécuter maintenant?", "XLApp - Mise à jour disponible", MessageBoxButtons.OKCancel);
                        if (!(DialogResult.OK == dr)) {
                            doUpdate = false;
                            return "Mise à jour annulée."; //
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
                            // ad.Update(); // Enlèvement SB
                            Uri DocPath = new Uri(Globals.ThisAddIn.Application.Path + "\\" + Globals.ThisAddIn.Application.Name); //test sb
                            Uri InstallerPath = new Uri("C:\\Program Files\\Common Files\\microsoft shared\\VSTO\\10.0\\VSTOINSTALLER.exe"); //test sb
                           // Uri RestarterPath = new Uri(CachePath + "WordRestarter.exe"); //enlève sb
                            Uri Updatelocation = new Uri(CurrentDep.UpdateLocation.ToString());

                            //Call VSTOInstaller Explicitely in "Silent Mode"
                            Process VstoInstallerProc = new System.Diagnostics.Process();
                            VstoInstallerProc.StartInfo.Arguments = " /S /I " + Updatelocation.AbsoluteUri;
                            VstoInstallerProc.StartInfo.FileName = InstallerPath.AbsoluteUri;
                            VstoInstallerProc.Start();

                            VstoInstallerProc.WaitForExit();
                            if (VstoInstallerProc.ExitCode == 0) {
                                string updatedVersDL = (DateTime.Now.Year % 100).ToString() + "." + CurrentDep.UpdatedVersion.ToString();
                                MessageBox.Show("La mise à jour de l'application à la version " + updatedVersDL + " a été réussi et sera effective au prochain redémarrage de l'application. Veuillez redémarrer l'application maintenant.", "Mise à jour - Version " + updatedVersDL);
                                return updatedVersDL;
                            } else {
                              //  MessageBox.Show("Échec de mise à jour: Exit Code (" + VstoInstallerProc.ExitCode.ToString() + ")");
                                return "Échec de mise à jour: Exit Code (" + VstoInstallerProc.ExitCode.ToString() + ")";
                            }
                               

                            //Call VSTOInstaller Explicitely in "Silent Mode"
                            // Process RestarterProc = new System.Diagnostics.Process();
                            // RestarterProc.StartInfo.Arguments = DocPath.AbsoluteUri;
                            // RestarterProc.StartInfo.FileName = RestarterPath.AbsoluteUri;
                            //  RestarterProc.Start();


                            //MessageBox.Show("The application has been upgraded, and will now restart.");
                            //MessageBox.Show("La mise à jour de l'application a été réussi et sera effective au prochain redémarrage.");
                            
                            //Application.Restart(); MODIF SB ENLÈVEMENT !
                        } catch (DeploymentDownloadException dde) {
                            //MessageBox.Show("Cannot install the latest version of the application. \n\nPlease check your network connection, or try again later. Error: " + dde);
                           // MessageBox.Show("Échec d'installation de la plus récente mise à jour. \n\nVeuillez vérifier votre connection, ou réessayer plus tard. Erreur: " + dde);
                            return "Échec d'installation de la plus récente mise à jour. \n\nVeuillez vérifier votre connection, ou réessayer plus tard. Erreur: " + dde;
                        }
                    }
                }
            }
            return "";
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
        //public string UpdatedVersion() {
        //    // How to get current the product version in C#?
        //    // Just give the reference to System.Deployment.Application and though it wont work in developement of the visual studio but it will work once the application is deployed.

        //    ////using System.Deployment.Application;
        //    ////using System.Reflection; 
        //    return ApplicationDeployment.IsNetworkDeployed
        //               ? ApplicationDeployment.CurrentDeployment.UpdatedVersion.ToString() // no. version après téléchargement.
        //               : Assembly.GetExecutingAssembly().GetName().Version.ToString(); //le 2e retourne : 1.0.0.0
        //    // si la version ici ne match pas celui indiqué dans le raccourci du desktop (et startmenu), on  copiera les dossiers -Projets, -Projets BackUp, ainsi que le fichier -Importation Bordereau du dossier Resources et le fichier -logo.png du dossier Images 
        //}
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
                customTaskPane = Globals.ThisAddIn.TaskPaneEstImposWPF;
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

                    return Globals.ThisAddIn.TaskPaneEstImposWPF;
                case "test2":

                    return Globals.ThisAddIn.TaskPaneEstImposWPF;
                default:
                    return Globals.ThisAddIn.TaskPaneEstImposWPF;
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

                //en attendant, la demande a été fait manuellement à l'utilisateur dans la méthod qui appelle le  InitiateFirstLaunch()

                List<int> oldExcelIDs = new List<int>();
                Process[] excelProcesses = Process.GetProcessesByName("Excel");
                foreach (Process pro in excelProcesses) { oldExcelIDs.Add(pro.Id); }

                Excel.Application xlApp;

                Excel.Workbook xlWorkBook;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                // http://stackoverflow.com/questions/1600502/programmatically-configuring-ms-words-trust-center-settings-using-c-sharp
                // a very simple way of opening an .xls file containing macros, without messing with the registry or Excel's trust settings. :
                xlApp.FileValidation = Microsoft.Office.Core.MsoFileValidationMode.msoFileValidationSkip;
                // fin a very simple way :)
                xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow; //26 mai 2017 https://msdn.microsoft.com/en-us/library/office/ff194819.aspx
                //Get the assembly information
                System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
                //Location is where the assembly is run from 
                string assemblyLocation = assemblyInfo.Location;
                //CodeBase is the location of the ClickOnce deployment files
                Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
                string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
                //xlApp.Caption = "";
                //                                                                   readonly=true
                xlWorkBook = xlApp.Workbooks.Open(ClickOnceLocation + "\\XLApp.dll", 0, true, 5, "False", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", false, false, 0, false, 1, 0);
                xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow; //26 mai 2017 https://msdn.microsoft.com/en-us/library/office/ff194819.aspx
                xlApp.Caption = "";
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                //xlApp.Visible = true
                //xlWorkBook = null;
                //xlApp = null;

                // updateDeskTopShortCutDescription("XLApp");
                System.Windows.MessageBox.Show("Pour utiliser l'application, veuillez lancer le raccourci par votre bureau ou par le menu démarrer.", "XLApp");


                Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible = false;
                AddInUtilities.UnConnectAddin();

                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
                xlWorkBook = null;
                xlApp = null;

                // https://stackoverflow.com/questions/17777545/closing-excel-application-process-in-c-sharp-after-data-access

                //--------Take the list of excel processes again and compare the IDs, if the Id is not in the old list is the one we just created, let's kill it!------
                excelProcesses = Process.GetProcessesByName("Excel");
                foreach (Process proc in excelProcesses) {
                    if (!oldExcelIDs.Contains(proc.Id)) {
                        try {
                            proc.Kill();
                        } catch {
                        }
                    }
                }

                return;
            }
            //si l'utilisateur ne veut pas démarrer l'app initialisation:
            Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible = false;
            AddInUtilities.UnConnectAddin();
            return;
        }

        internal static void ShowParamProjet() {
            Globals.ThisAddIn.Application.Run("ShowParamProjet");
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
            xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow; //26 mai 2017 https://msdn.microsoft.com/en-us/library/office/ff194819.aspx

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
            xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow; //26 mai 2017 https://msdn.microsoft.com/en-us/library/office/ff194819.aspx

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
            return Globals.ThisAddIn.TaskPaneInterfaceVert.Visible;
        }
        public void AdjustComboBoxLine(string indText)
        {
            ManageTaskPaneRibbon.AdjustComboBoxLine(indText);
        }
        public void ShowAppVertBar() //testing 11-12-2016 from vba
        {
            Globals.ThisAddIn.TaskPaneInterfaceVert.Visible = !(Globals.ThisAddIn.TaskPaneInterfaceVert.Visible);
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
                isMyApp = (Globals.ThisAddIn.Application.Caption.IndexOf("LGSL+") > -1);
                return (Globals.ThisAddIn.Application.ActiveWorkbook.IsAddin);
            }
             catch
             {
                // Si erreur, c'est parce que c'est un addin (par essai-erreur)
                isMyApp = (Globals.ThisAddIn.Application.Caption.IndexOf("LGSL+") > -1);
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
        internal static void ImportSoum() 
        {
            Globals.ThisAddIn.Application.Run("importBordereauExistant");
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
