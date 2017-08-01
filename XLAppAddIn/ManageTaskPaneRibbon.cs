using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using System.IO;
using IWshRuntimeLibrary;
using System.Deployment.Application;
using System.Reflection;

namespace XLAppAddIn
{

    public partial class ManageTaskPaneRibbon
    {
        //MyUserControl MyUserControl1 = new MyUserControl();

        private void ManageTaskPaneRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //set properties : 
            //this.groupDeplacement.Visible = new EventHandler(saisieTabs_setVisible);
            //this.groupNavigation.Visible = new EventHandler(saisieTabs_setVisible);
            //this.groupSaisieItems.Visible = new EventHandler(saisieTabs_setVisible);
            
        }
        //public bool saisieTabs_setVisible(object sender, EventArgs e)
        //{

        //    return false;

        //}
        public void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            bool isMyApp;
            if (AddInUtilities.GetIsAddIn(out isMyApp)) //UNCHECKButton
            {
                //ShowXLBackStageView:
                //Globals.Ribbons.CustomRibbon.Tabs[Your tab id].RibbonUI.ActivateTab("");
                //Globals.Ribbons.ManageTaskPaneRibbon.RibbonUI.ActivateTab("FileTab");
                //Excel 2010 or higher: Build in way to activate tab
                if (Globals.Ribbons.ManageTaskPaneRibbon.RibbonUI != null)
                {
                    toggleButton1.Checked = false;
                    System.Windows.Forms.SendKeys.Send("%{f}%"); //va aller dans le backstage view de Excel
                    
                    return;
                    //Globals.Ribbons.ManageTaskPaneRibbon.RibbonUI.ActivateTab("TabHome");
                }
            }
            if (!isMyApp)
            {
                //Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible = false; // messagebox pour avertir l'utilisateur ou fermer la visibilité...

                string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                if (!System.IO.File.Exists(userProfile + "\\Desktop\\" + "XLApp" + ".lnk")) {
                    // code si aucun raccourci sur le bureau :
                    string msgErr = "";

                    // vérifier si c'est au moins excel 2013 et 64 bit (version 15)
                    int noVers = int.Parse(Globals.ThisAddIn.Application.Version.ToString().Split('.')[0]);
                    //bool Is64bit = Environment.GetEnvironmentVariable("ProgramW6432").Length > 0

                    if (noVers < 15) {
                        msgErr = "Pour finaliser l'installation, la version d'Excel 2013 ou plus récente est requise, option 64 bit.";
                    }


                    //vérifier si l'accès à la sécurité est activé avant de poursuivre, sinon, informez l'utilisateur comment le faire.
                    try {
                        var VDP = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                        if (VDP != null) VDP = null;
                    } catch {
                        if (msgErr != "")
                            msgErr += "\n\nEnsuite, vous devez configurer une option de sécurité dans Excel en suivant cette procédure :\n\nFichier > Options > Paramètres du Centre de gestion de la confidentialité > Paramètres des macros > Cocher \"Accès approuv au modèle d'objet du projet VBA\"";
                        else
                             msgErr += "Pour finaliser l'installation, configurer une option de sécurité en suivant cette procédure :\n\nFichier > Options > Paramètres du Centre de gestion de la confidentialité > Paramètres des macros > Cocher \"Accès approuv au modèle d'objet du projet VBA\"";
                    }

                    if (msgErr != "") {
                        System.Windows.MessageBox.Show(msgErr, "XLApp");
                        toggleButton1.Checked = false;
                        return;
                    }


                    AddInUtilities.InitiateFirstLaunch();


                    //    if (AddInUtilities.InitiateFirstLaunch()) {
                    //        updateDeskTopShortCutDescription("XLApp");
                    //        System.Windows.MessageBox.Show("Pour utiliser l'application, veuillez lancer le raccourci par votre bureau ou par le menu démarrer.", "XLApp");
                    //    }

                    //Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible = false;
                    //AddInUtilities.UnConnectAddin();
                }
                else {
                    DialogResult result1 = MessageBox.Show("Lancer l'application?",
                                        "XLApp",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question,
                                        MessageBoxDefaultButton.Button1);


                    if (result1 == DialogResult.Yes)
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible = false;
                        AddInUtilities.UnConnectAddin();
                        AddInUtilities.LaunchApp();
                    }
                    else
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible = false;
                        AddInUtilities.UnConnectAddin();
                    }

                }
                return;
            }

        Globals.ThisAddIn.TaskPaneInterfaceVert.Visible = ((RibbonToggleButton)sender).Checked;
            


            //Globals.ThisAddIn.Application.Run("resizeWindow");

                //if (!Globals.Application.ActiveWorkbook.IsAddin)
                //    Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;
                //else
                //     Globals.Application.Run("SheetList_RDB");

        }

        private static void updateDeskTopShortCutDescription(string shortcutName) {
            WshShell wsh = new WshShell();
            IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + shortcutName + ".lnk") as IWshRuntimeLibrary.IWshShortcut;
            //shortcut.Arguments = "";
            //shortcut.TargetPath = "c:\\app\\myftp.exe";

            string curVersion = ApplicationDeployment.IsNetworkDeployed ? ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString() // retourne la bonne version en exécution
                                : Assembly.GetExecutingAssembly().GetName().Version.ToString(); //le 2e retourne : 1.0.0.0
            shortcut.Description = "XLApp, Version " + DateTime.Now.Year % 100 + "." + curVersion;
            //shortcut.WorkingDirectory = "c:\\app";
            //shortcut.IconLocation = "specify icon location";
            shortcut.Save();
        }

        //ToggleCopyPaste
        public static void ToggleCopyPasteRibbon(string enabled)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.buttonPaste.Enabled = Convert.ToBoolean(enabled);
            Globals.Ribbons.ManageTaskPaneRibbon.buttonClearClip.Enabled = Convert.ToBoolean(enabled);
        }

        public static void ShowRibbonAddinTab()
        {
            if (!(Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible)) Globals.Ribbons.ManageTaskPaneRibbon.tab2.Visible = true;

        }

        public static void ShowOrHideUserControl()  // static allows it to be called from other class
        {
            if (Globals.ThisAddIn.TaskPaneInterfaceVert.Visible == true)
                Globals.ThisAddIn.TaskPaneInterfaceVert.Visible = false;
            else
                Globals.ThisAddIn.TaskPaneInterfaceVert.Visible = true;
        }

        public static void HideUserControl()  // static allows it to be called from other class
        {

            Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = false;
            Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
            Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
            Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = false;
            Globals.Ribbons.ManageTaskPaneRibbon.groupRessources.Visible = false;
            
            Globals.ThisAddIn.TaskPaneInterfaceVert.Visible = false;
        }

        public static void ShowUserControl()  // static allows it to be called from other class
        {
                Globals.ThisAddIn.TaskPaneInterfaceVert.Visible = true;


        }
        public static int GetUserControlWidth()  // static allows it to be called from other class
        {
            return Globals.ThisAddIn.TaskPaneInterfaceVert.Width;

        }
        // set ribbon text caused by a change in VBA
        public static void AdjustComboBoxLine(string indText)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.comboBoxLignes.Text = indText;
        }
       //GROUPE Saisie des items
        private void buttonArt_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.insertArticle();
        }

        private void buttonArtAss_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.insertArticleAss();
        }

        private void buttonAss_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.insertAss();
        }

        private void buttonOrder_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.OrderArt();
        }

        private void buttonFormula_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.ShowFunctions();
        }

        private void comboBoxLignes_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // changer celui de VBA :
            string cSharpIndex = comboBoxLignes.Text;
            AddInUtilities.TextChangedComboBoxLignes(cSharpIndex);
        }

        private void buttonLignes_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.InsertLineSaisieBanque();
        }

        private void buttonSupLigne_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.SupprimerLigne();
        }

        private void buttonSupProd_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.SupprimerProduits();
        }

        private void buttonSupArt_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.SupprimerLignesMultiples();  // supprimer 1 Article
        }

        private void buttonRefr_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.Macroretablirarticles();
        }
           //navigation
        private void buttonNavFirst_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.GotoPremier();
        }

        private void buttonNavPrev_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.GotoPrec();
        }

        private void buttonNavNext_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.GotoSuiv();
        }

        private void buttonNavLast_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.GotoLast();
        }

        private void buttonAllerArt_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.GotoArticle();
        }

        private void buttonParam_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.MODART();
        }


        private void buttonImport_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.ShowBordImp();
        }

        private void toggleButtonVerif_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPaneVerifProjet.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void buttonNouvP_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.NouveauProjet();
        }

        private void buttonOuvrirP_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.OuvrirProjet();
        }

        private void buttonEnregP_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.EnregProj();
        }

        private void buttonEnregSous_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.EnregProjSous();
        }

        private void buttonFermerProjet_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.FermerProjet();
            // then toggle needed to make ribbon groups invisible except "Application Excel 0365"
            //if (AddInUtilities.GetIsAddIn())
            //{
            //    Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = false;
            //    Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
            //    Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
            //    Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = false;
            //    Globals.ThisAddIn.TaskPane.Visible = false;
            //}

        }

        private void buttonPaste_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.Coller();
        }

        private void buttonCopy_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.Copier();
        }

        private void buttonCut_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.Couper();
        }

        private void buttonClearClip_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.ClearClip();
        }

        private void AtoZ_Click(object sender, RibbonControlEventArgs e)
        {
            var ribButton = sender as RibbonButton;
            if (ribButton != null)
            {
                AddInUtilities.FilterProduitsByLetterAndGoToFirstLetter(ribButton.Tag.ToString());
            }
        }

        private void groupDeplacement_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.DisplayClipboardWindow = !Globals.ThisAddIn.Application.DisplayClipboardWindow;
        }

        private void buttonProdRess_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.ListerRessources();
        }

        private void buttonImp_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.IMPRBORD();
        }

        private void buttonSignAuto_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.SignatureAuto();
        }

        private void buttonZImpAuto_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.ZoneImpAuto();
        }



        private void buttonAffMsqUnit_Click(object sender, RibbonControlEventArgs e)
        {
            AddInUtilities.AMCoutU();
        }

        private void buttonParamProjet_Click(object sender, RibbonControlEventArgs e) {
            AddInUtilities.ShowParamProjet();
        }

        private void buttonImportSoum_Click(object sender, RibbonControlEventArgs e) {
            AddInUtilities.ImportSoum();
        }

        private void buttonZImpAutoRessP_Click(object sender, RibbonControlEventArgs e) {
            AddInUtilities.ZoneImpAuto();
        }


        private void buttonIconRessP_Click(object sender, RibbonControlEventArgs e) {
            AddInUtilities.SelectIconForRapp();
        }

        private void buttonIconB_Click(object sender, RibbonControlEventArgs e) {
            AddInUtilities.SelectIconForRapp();
        }
    }
}
