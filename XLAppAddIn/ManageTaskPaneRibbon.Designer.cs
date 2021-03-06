﻿namespace XLAppAddIn
{
    partial class ManageTaskPaneRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ManageTaskPaneRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl15 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl16 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl3 = this.Factory.CreateRibbonDialogLauncher();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ManageTaskPaneRibbon));
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.groupGestProjet = this.Factory.CreateRibbonGroup();
            this.groupDeplacement = this.Factory.CreateRibbonGroup();
            this.groupNavigation = this.Factory.CreateRibbonGroup();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.groupSaisieItems = this.Factory.CreateRibbonGroup();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.comboBoxLignes = this.Factory.CreateRibbonComboBox();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.box2 = this.Factory.CreateRibbonBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.groupRessources = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.buttonGroup3 = this.Factory.CreateRibbonButtonGroup();
            this.separator7 = this.Factory.CreateRibbonSeparator();
            this.groupRessProj = this.Factory.CreateRibbonGroup();
            this.separator10 = this.Factory.CreateRibbonSeparator();
            this.labelBogus = this.Factory.CreateRibbonLabel();
            this.groupBordereau = this.Factory.CreateRibbonGroup();
            this.separator9 = this.Factory.CreateRibbonSeparator();
            this.separator8 = this.Factory.CreateRibbonSeparator();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.buttonNouvP = this.Factory.CreateRibbonButton();
            this.buttonOuvrirP = this.Factory.CreateRibbonButton();
            this.buttonEnregP = this.Factory.CreateRibbonButton();
            this.buttonFermerProjet = this.Factory.CreateRibbonButton();
            this.buttonParamProjet = this.Factory.CreateRibbonButton();
            this.buttonEnregSous = this.Factory.CreateRibbonButton();
            this.buttonPaste = this.Factory.CreateRibbonButton();
            this.buttonCopy = this.Factory.CreateRibbonButton();
            this.buttonCut = this.Factory.CreateRibbonButton();
            this.buttonClearClip = this.Factory.CreateRibbonButton();
            this.buttonNavFirst = this.Factory.CreateRibbonButton();
            this.buttonNavPrev = this.Factory.CreateRibbonButton();
            this.buttonNavNext = this.Factory.CreateRibbonButton();
            this.buttonNavLast = this.Factory.CreateRibbonButton();
            this.buttonAllerArt = this.Factory.CreateRibbonButton();
            this.buttonParam = this.Factory.CreateRibbonButton();
            this.buttonOrder = this.Factory.CreateRibbonButton();
            this.galleryImport = this.Factory.CreateRibbonGallery();
            this.buttonImportExcel = this.Factory.CreateRibbonButton();
            this.buttonImportAccess = this.Factory.CreateRibbonButton();
            this.buttonPressePapier = this.Factory.CreateRibbonButton();
            this.buttonArt = this.Factory.CreateRibbonButton();
            this.buttonArtAss = this.Factory.CreateRibbonButton();
            this.buttonAss = this.Factory.CreateRibbonButton();
            this.buttonLignes = this.Factory.CreateRibbonButton();
            this.buttonFormula = this.Factory.CreateRibbonButton();
            this.buttonSupLigne = this.Factory.CreateRibbonButton();
            this.buttonSupProd = this.Factory.CreateRibbonButton();
            this.buttonSupArt = this.Factory.CreateRibbonButton();
            this.buttonRefr = this.Factory.CreateRibbonButton();
            this.toggleButtonVerif = this.Factory.CreateRibbonToggleButton();
            this.buttonA = this.Factory.CreateRibbonButton();
            this.buttonB = this.Factory.CreateRibbonButton();
            this.buttonC = this.Factory.CreateRibbonButton();
            this.buttonD = this.Factory.CreateRibbonButton();
            this.buttonE = this.Factory.CreateRibbonButton();
            this.buttonF = this.Factory.CreateRibbonButton();
            this.buttonG = this.Factory.CreateRibbonButton();
            this.buttonH = this.Factory.CreateRibbonButton();
            this.buttonI = this.Factory.CreateRibbonButton();
            this.buttonJ = this.Factory.CreateRibbonButton();
            this.buttonK = this.Factory.CreateRibbonButton();
            this.buttonL = this.Factory.CreateRibbonButton();
            this.buttonM = this.Factory.CreateRibbonButton();
            this.buttonN = this.Factory.CreateRibbonButton();
            this.buttonO = this.Factory.CreateRibbonButton();
            this.buttonP = this.Factory.CreateRibbonButton();
            this.buttonQ = this.Factory.CreateRibbonButton();
            this.buttonR = this.Factory.CreateRibbonButton();
            this.buttonS = this.Factory.CreateRibbonButton();
            this.buttonT = this.Factory.CreateRibbonButton();
            this.buttonU = this.Factory.CreateRibbonButton();
            this.buttonV = this.Factory.CreateRibbonButton();
            this.buttonW = this.Factory.CreateRibbonButton();
            this.buttonX = this.Factory.CreateRibbonButton();
            this.buttonY = this.Factory.CreateRibbonButton();
            this.buttonZ = this.Factory.CreateRibbonButton();
            this.button0 = this.Factory.CreateRibbonButton();
            this.buttonRefrRess = this.Factory.CreateRibbonButton();
            this.buttonProdRess = this.Factory.CreateRibbonButton();
            this.buttonZImpAutoRessP = this.Factory.CreateRibbonButton();
            this.buttonIconRessP = this.Factory.CreateRibbonButton();
            this.buttonImportSoum = this.Factory.CreateRibbonButton();
            this.buttonImp = this.Factory.CreateRibbonButton();
            this.buttonSignAuto = this.Factory.CreateRibbonButton();
            this.buttonZImpAuto = this.Factory.CreateRibbonButton();
            this.buttonIconB = this.Factory.CreateRibbonButton();
            this.buttonAffMsqUnit = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupGestProjet.SuspendLayout();
            this.groupDeplacement.SuspendLayout();
            this.groupNavigation.SuspendLayout();
            this.groupSaisieItems.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.groupRessources.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.buttonGroup3.SuspendLayout();
            this.groupRessProj.SuspendLayout();
            this.groupBordereau.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.groupGestProjet);
            this.tab2.Groups.Add(this.groupDeplacement);
            this.tab2.Groups.Add(this.groupNavigation);
            this.tab2.Groups.Add(this.groupSaisieItems);
            this.tab2.Groups.Add(this.groupRessources);
            this.tab2.Groups.Add(this.groupRessProj);
            this.tab2.Groups.Add(this.groupBordereau);
            this.tab2.Label = "XLApp";
            this.tab2.Name = "tab2";
            this.tab2.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabAddIns");
            // 
            // group1
            // 
            ribbonDialogLauncherImpl1.Visible = false;
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.toggleButton1);
            this.group1.Label = "  Application Excel O365";
            this.group1.Name = "group1";
            // 
            // groupGestProjet
            // 
            this.groupGestProjet.Items.Add(this.buttonNouvP);
            this.groupGestProjet.Items.Add(this.buttonOuvrirP);
            this.groupGestProjet.Items.Add(this.buttonEnregP);
            this.groupGestProjet.Items.Add(this.buttonFermerProjet);
            this.groupGestProjet.Items.Add(this.buttonParamProjet);
            this.groupGestProjet.Items.Add(this.buttonEnregSous);
            this.groupGestProjet.Label = "Gestionnaire de projets";
            this.groupGestProjet.Name = "groupGestProjet";
            this.groupGestProjet.Visible = false;
            // 
            // groupDeplacement
            // 
            this.groupDeplacement.DialogLauncher = ribbonDialogLauncherImpl2;
            this.groupDeplacement.Items.Add(this.buttonPaste);
            this.groupDeplacement.Items.Add(this.buttonCopy);
            this.groupDeplacement.Items.Add(this.buttonCut);
            this.groupDeplacement.Items.Add(this.buttonClearClip);
            this.groupDeplacement.Label = "Déplacement";
            this.groupDeplacement.Name = "groupDeplacement";
            this.groupDeplacement.Visible = false;
            this.groupDeplacement.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.groupDeplacement_DialogLauncherClick);
            // 
            // groupNavigation
            // 
            this.groupNavigation.Items.Add(this.buttonNavFirst);
            this.groupNavigation.Items.Add(this.buttonNavPrev);
            this.groupNavigation.Items.Add(this.buttonNavNext);
            this.groupNavigation.Items.Add(this.buttonNavLast);
            this.groupNavigation.Items.Add(this.separator4);
            this.groupNavigation.Items.Add(this.buttonAllerArt);
            this.groupNavigation.Items.Add(this.buttonParam);
            this.groupNavigation.Items.Add(this.buttonOrder);
            this.groupNavigation.Label = "Navigation";
            this.groupNavigation.Name = "groupNavigation";
            this.groupNavigation.Visible = false;
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // groupSaisieItems
            // 
            this.groupSaisieItems.Items.Add(this.galleryImport);
            this.groupSaisieItems.Items.Add(this.separator3);
            this.groupSaisieItems.Items.Add(this.buttonArt);
            this.groupSaisieItems.Items.Add(this.buttonArtAss);
            this.groupSaisieItems.Items.Add(this.buttonAss);
            this.groupSaisieItems.Items.Add(this.separator1);
            this.groupSaisieItems.Items.Add(this.box1);
            this.groupSaisieItems.Items.Add(this.separator5);
            this.groupSaisieItems.Items.Add(this.box2);
            this.groupSaisieItems.Items.Add(this.separator2);
            this.groupSaisieItems.Items.Add(this.buttonRefr);
            this.groupSaisieItems.Items.Add(this.separator6);
            this.groupSaisieItems.Items.Add(this.toggleButtonVerif);
            this.groupSaisieItems.Label = "Saisie d\'items";
            this.groupSaisieItems.Name = "groupSaisieItems";
            this.groupSaisieItems.Visible = false;
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.buttonLignes);
            this.box1.Items.Add(this.comboBoxLignes);
            this.box1.Items.Add(this.buttonFormula);
            this.box1.Name = "box1";
            // 
            // comboBoxLignes
            // 
            ribbonDropDownItemImpl1.Label = "+ 1 ligne";
            ribbonDropDownItemImpl2.Label = "+ 2";
            ribbonDropDownItemImpl3.Label = "+ 3";
            ribbonDropDownItemImpl4.Label = "+ 4";
            ribbonDropDownItemImpl5.Label = "+ 5";
            ribbonDropDownItemImpl6.Label = "+ 6";
            ribbonDropDownItemImpl7.Label = "+ 7";
            ribbonDropDownItemImpl8.Label = "+ 8";
            ribbonDropDownItemImpl9.Label = "+ 9";
            ribbonDropDownItemImpl10.Label = "+ 10";
            ribbonDropDownItemImpl11.Label = "+ 11";
            ribbonDropDownItemImpl12.Label = "+ 12";
            ribbonDropDownItemImpl13.Label = "+ 13";
            ribbonDropDownItemImpl14.Label = "+ 14";
            ribbonDropDownItemImpl15.Label = "+ 15";
            ribbonDropDownItemImpl16.Label = "+ 16";
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl1);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl2);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl3);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl4);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl5);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl6);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl7);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl8);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl9);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl10);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl11);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl12);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl13);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl14);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl15);
            this.comboBoxLignes.Items.Add(ribbonDropDownItemImpl16);
            this.comboBoxLignes.Label = " ";
            this.comboBoxLignes.MaxLength = 16;
            this.comboBoxLignes.Name = "comboBoxLignes";
            this.comboBoxLignes.ShowImage = true;
            this.comboBoxLignes.SizeString = "+ 1 ligne";
            this.comboBoxLignes.Text = "+ 1 ligne";
            this.comboBoxLignes.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxLignes_TextChanged);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.buttonSupLigne);
            this.box2.Items.Add(this.buttonSupProd);
            this.box2.Items.Add(this.buttonSupArt);
            this.box2.Name = "box2";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // groupRessources
            // 
            this.groupRessources.Items.Add(this.buttonGroup1);
            this.groupRessources.Items.Add(this.buttonGroup2);
            this.groupRessources.Items.Add(this.buttonGroup3);
            this.groupRessources.Items.Add(this.separator7);
            this.groupRessources.Items.Add(this.buttonRefrRess);
            this.groupRessources.Label = "Saisie de ressources";
            this.groupRessources.Name = "groupRessources";
            this.groupRessources.Visible = false;
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.buttonA);
            this.buttonGroup1.Items.Add(this.buttonB);
            this.buttonGroup1.Items.Add(this.buttonC);
            this.buttonGroup1.Items.Add(this.buttonD);
            this.buttonGroup1.Items.Add(this.buttonE);
            this.buttonGroup1.Items.Add(this.buttonF);
            this.buttonGroup1.Items.Add(this.buttonG);
            this.buttonGroup1.Items.Add(this.buttonH);
            this.buttonGroup1.Items.Add(this.buttonI);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.buttonJ);
            this.buttonGroup2.Items.Add(this.buttonK);
            this.buttonGroup2.Items.Add(this.buttonL);
            this.buttonGroup2.Items.Add(this.buttonM);
            this.buttonGroup2.Items.Add(this.buttonN);
            this.buttonGroup2.Items.Add(this.buttonO);
            this.buttonGroup2.Items.Add(this.buttonP);
            this.buttonGroup2.Items.Add(this.buttonQ);
            this.buttonGroup2.Items.Add(this.buttonR);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // buttonGroup3
            // 
            this.buttonGroup3.Items.Add(this.buttonS);
            this.buttonGroup3.Items.Add(this.buttonT);
            this.buttonGroup3.Items.Add(this.buttonU);
            this.buttonGroup3.Items.Add(this.buttonV);
            this.buttonGroup3.Items.Add(this.buttonW);
            this.buttonGroup3.Items.Add(this.buttonX);
            this.buttonGroup3.Items.Add(this.buttonY);
            this.buttonGroup3.Items.Add(this.buttonZ);
            this.buttonGroup3.Items.Add(this.button0);
            this.buttonGroup3.Name = "buttonGroup3";
            // 
            // separator7
            // 
            this.separator7.Name = "separator7";
            // 
            // groupRessProj
            // 
            this.groupRessProj.DialogLauncher = ribbonDialogLauncherImpl3;
            this.groupRessProj.Items.Add(this.buttonProdRess);
            this.groupRessProj.Items.Add(this.separator10);
            this.groupRessProj.Items.Add(this.labelBogus);
            this.groupRessProj.Items.Add(this.buttonZImpAutoRessP);
            this.groupRessProj.Items.Add(this.buttonIconRessP);
            this.groupRessProj.Label = "Ressources de projet";
            this.groupRessProj.Name = "groupRessProj";
            this.groupRessProj.Visible = false;
            // 
            // separator10
            // 
            this.separator10.Name = "separator10";
            // 
            // labelBogus
            // 
            this.labelBogus.Label = " ";
            this.labelBogus.Name = "labelBogus";
            // 
            // groupBordereau
            // 
            this.groupBordereau.Items.Add(this.buttonImportSoum);
            this.groupBordereau.Items.Add(this.separator9);
            this.groupBordereau.Items.Add(this.buttonImp);
            this.groupBordereau.Items.Add(this.buttonSignAuto);
            this.groupBordereau.Items.Add(this.buttonZImpAuto);
            this.groupBordereau.Items.Add(this.buttonIconB);
            this.groupBordereau.Items.Add(this.separator8);
            this.groupBordereau.Items.Add(this.buttonAffMsqUnit);
            this.groupBordereau.Label = "Configuration de bordereau";
            this.groupBordereau.Name = "groupBordereau";
            this.groupBordereau.Visible = false;
            // 
            // separator9
            // 
            this.separator9.Name = "separator9";
            // 
            // separator8
            // 
            this.separator8.Name = "separator8";
            // 
            // toggleButton1
            // 
            this.toggleButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton1.Image = ((System.Drawing.Image)(resources.GetObject("toggleButton1.Image")));
            this.toggleButton1.Label = "Barre XLApp";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.OfficeImageId = "TaskPanesMenu";
            this.toggleButton1.ScreenTip = "Afficher/Masquer la barre verticale";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.SuperTip = "Afficher/Masquer la barre des tâches vertical";
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // buttonNouvP
            // 
            this.buttonNouvP.Label = "Nouveau";
            this.buttonNouvP.Name = "buttonNouvP";
            this.buttonNouvP.OfficeImageId = "NewXmlPage";
            this.buttonNouvP.ShowImage = true;
            this.buttonNouvP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNouvP_Click);
            // 
            // buttonOuvrirP
            // 
            this.buttonOuvrirP.Label = "Ouvrir";
            this.buttonOuvrirP.Name = "buttonOuvrirP";
            this.buttonOuvrirP.OfficeImageId = "OpenFolder";
            this.buttonOuvrirP.ShowImage = true;
            this.buttonOuvrirP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonOuvrirP_Click);
            // 
            // buttonEnregP
            // 
            this.buttonEnregP.Label = "Enregistrer";
            this.buttonEnregP.Name = "buttonEnregP";
            this.buttonEnregP.OfficeImageId = "FileSave";
            this.buttonEnregP.ShowImage = true;
            this.buttonEnregP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEnregP_Click);
            // 
            // buttonFermerProjet
            // 
            this.buttonFermerProjet.Label = "Fermer Projet";
            this.buttonFermerProjet.Name = "buttonFermerProjet";
            this.buttonFermerProjet.OfficeImageId = "CloseAllPages";
            this.buttonFermerProjet.ShowImage = true;
            this.buttonFermerProjet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFermerProjet_Click);
            // 
            // buttonParamProjet
            // 
            this.buttonParamProjet.Label = "Paramètres";
            this.buttonParamProjet.Name = "buttonParamProjet";
            this.buttonParamProjet.OfficeImageId = "TablesResetToDefault";
            this.buttonParamProjet.ShowImage = true;
            this.buttonParamProjet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonParamProjet_Click);
            // 
            // buttonEnregSous
            // 
            this.buttonEnregSous.Label = "Enregistrer sous";
            this.buttonEnregSous.Name = "buttonEnregSous";
            this.buttonEnregSous.OfficeImageId = "FileSaveAs";
            this.buttonEnregSous.ShowImage = true;
            this.buttonEnregSous.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEnregSous_Click);
            // 
            // buttonPaste
            // 
            this.buttonPaste.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonPaste.Enabled = false;
            this.buttonPaste.Label = "Coller";
            this.buttonPaste.Name = "buttonPaste";
            this.buttonPaste.OfficeImageId = "Paste";
            this.buttonPaste.ShowImage = true;
            this.buttonPaste.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonPaste_Click);
            // 
            // buttonCopy
            // 
            this.buttonCopy.Label = "Copier";
            this.buttonCopy.Name = "buttonCopy";
            this.buttonCopy.OfficeImageId = "Copy";
            this.buttonCopy.ShowImage = true;
            this.buttonCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCopy_Click);
            // 
            // buttonCut
            // 
            this.buttonCut.Label = "Couper";
            this.buttonCut.Name = "buttonCut";
            this.buttonCut.OfficeImageId = "Cut";
            this.buttonCut.ShowImage = true;
            this.buttonCut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCut_Click);
            // 
            // buttonClearClip
            // 
            this.buttonClearClip.Enabled = false;
            this.buttonClearClip.Label = "Réinitialiser";
            this.buttonClearClip.Name = "buttonClearClip";
            this.buttonClearClip.OfficeImageId = "PasteInk";
            this.buttonClearClip.ShowImage = true;
            this.buttonClearClip.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonClearClip_Click);
            // 
            // buttonNavFirst
            // 
            this.buttonNavFirst.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonNavFirst.Label = "Premier";
            this.buttonNavFirst.Name = "buttonNavFirst";
            this.buttonNavFirst.OfficeImageId = "GoRtlDown";
            this.buttonNavFirst.ShowImage = true;
            this.buttonNavFirst.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNavFirst_Click);
            // 
            // buttonNavPrev
            // 
            this.buttonNavPrev.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonNavPrev.Label = "Précède";
            this.buttonNavPrev.Name = "buttonNavPrev";
            this.buttonNavPrev.OfficeImageId = "ScreenNavigatorBack";
            this.buttonNavPrev.ShowImage = true;
            this.buttonNavPrev.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNavPrev_Click);
            // 
            // buttonNavNext
            // 
            this.buttonNavNext.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonNavNext.Label = "Suivant";
            this.buttonNavNext.Name = "buttonNavNext";
            this.buttonNavNext.OfficeImageId = "ScreenNavigatorForward";
            this.buttonNavNext.ShowImage = true;
            this.buttonNavNext.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNavNext_Click);
            // 
            // buttonNavLast
            // 
            this.buttonNavLast.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonNavLast.Label = "Dernier";
            this.buttonNavLast.Name = "buttonNavLast";
            this.buttonNavLast.OfficeImageId = "GoLtrDown";
            this.buttonNavLast.ShowImage = true;
            this.buttonNavLast.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNavLast_Click);
            // 
            // buttonAllerArt
            // 
            this.buttonAllerArt.Label = "Aller à Article";
            this.buttonAllerArt.Name = "buttonAllerArt";
            this.buttonAllerArt.OfficeImageId = "GanttNext";
            this.buttonAllerArt.ScreenTip = "Aller à l\'Article";
            this.buttonAllerArt.ShowImage = true;
            this.buttonAllerArt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAllerArt_Click);
            // 
            // buttonParam
            // 
            this.buttonParam.Label = "Paramètres";
            this.buttonParam.Name = "buttonParam";
            this.buttonParam.OfficeImageId = "GroupChartProperties";
            this.buttonParam.ScreenTip = "Paramètres d\'Articles";
            this.buttonParam.ShowImage = true;
            this.buttonParam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonParam_Click);
            // 
            // buttonOrder
            // 
            this.buttonOrder.Label = "Ordonner";
            this.buttonOrder.Name = "buttonOrder";
            this.buttonOrder.OfficeImageId = "GroupCalendarOptions";
            this.buttonOrder.ScreenTip = "Ordonner Articles";
            this.buttonOrder.ShowImage = true;
            this.buttonOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonOrder_Click);
            // 
            // galleryImport
            // 
            this.galleryImport.Buttons.Add(this.buttonImportExcel);
            this.galleryImport.Buttons.Add(this.buttonImportAccess);
            this.galleryImport.Buttons.Add(this.buttonPressePapier);
            this.galleryImport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.galleryImport.Label = "Importer Bordereau";
            this.galleryImport.Name = "galleryImport";
            this.galleryImport.OfficeImageId = "Import";
            this.galleryImport.ShowImage = true;
            // 
            // buttonImportExcel
            // 
            this.buttonImportExcel.Label = "Format Excel";
            this.buttonImportExcel.Name = "buttonImportExcel";
            this.buttonImportExcel.OfficeImageId = "ImportExcel";
            this.buttonImportExcel.ShowImage = true;
            this.buttonImportExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImport_Click);
            // 
            // buttonImportAccess
            // 
            this.buttonImportAccess.Label = "Format Access";
            this.buttonImportAccess.Name = "buttonImportAccess";
            this.buttonImportAccess.OfficeImageId = "ImportAccess";
            this.buttonImportAccess.ShowImage = true;
            this.buttonImportAccess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImport_Click);
            // 
            // buttonPressePapier
            // 
            this.buttonPressePapier.Label = "Presse-papiers";
            this.buttonPressePapier.Name = "buttonPressePapier";
            this.buttonPressePapier.OfficeImageId = "ShowClipboard";
            this.buttonPressePapier.ShowImage = true;
            this.buttonPressePapier.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImport_Click);
            // 
            // buttonArt
            // 
            this.buttonArt.Label = "Insérer Article";
            this.buttonArt.Name = "buttonArt";
            this.buttonArt.OfficeImageId = "InsertRow";
            this.buttonArt.ShowImage = true;
            this.buttonArt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonArt_Click);
            // 
            // buttonArtAss
            // 
            this.buttonArtAss.Label = "Articles Assemblés";
            this.buttonArtAss.Name = "buttonArtAss";
            this.buttonArtAss.OfficeImageId = "Bullets";
            this.buttonArtAss.ShowImage = true;
            this.buttonArtAss.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonArtAss_Click);
            // 
            // buttonAss
            // 
            this.buttonAss.Label = "Assemblages";
            this.buttonAss.Name = "buttonAss";
            this.buttonAss.OfficeImageId = "PivotPlusMinusFieldHeadersShowHide";
            this.buttonAss.ShowImage = true;
            this.buttonAss.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAss_Click);
            // 
            // buttonLignes
            // 
            this.buttonLignes.Label = "Insérer Lignes";
            this.buttonLignes.Name = "buttonLignes";
            this.buttonLignes.OfficeImageId = "TableRowsInsertAboveExcel";
            this.buttonLignes.ShowImage = true;
            this.buttonLignes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLignes_Click);
            // 
            // buttonFormula
            // 
            this.buttonFormula.Label = "Fonction XLApp";
            this.buttonFormula.Name = "buttonFormula";
            this.buttonFormula.OfficeImageId = "EditFormula";
            this.buttonFormula.ShowImage = true;
            this.buttonFormula.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFormula_Click);
            // 
            // buttonSupLigne
            // 
            this.buttonSupLigne.Label = "Ligne(s)";
            this.buttonSupLigne.Name = "buttonSupLigne";
            this.buttonSupLigne.OfficeImageId = "SheetRowsDelete";
            this.buttonSupLigne.ShowImage = true;
            this.buttonSupLigne.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSupLigne_Click);
            // 
            // buttonSupProd
            // 
            this.buttonSupProd.Label = "Produit(s)";
            this.buttonSupProd.Name = "buttonSupProd";
            this.buttonSupProd.OfficeImageId = "CellsDelete";
            this.buttonSupProd.ShowImage = true;
            this.buttonSupProd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSupProd_Click);
            // 
            // buttonSupArt
            // 
            this.buttonSupArt.Label = "Article(s)";
            this.buttonSupArt.Name = "buttonSupArt";
            this.buttonSupArt.OfficeImageId = "Delete";
            this.buttonSupArt.ShowImage = true;
            this.buttonSupArt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSupArt_Click);
            // 
            // buttonRefr
            // 
            this.buttonRefr.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRefr.Label = "Rafraîchir";
            this.buttonRefr.Name = "buttonRefr";
            this.buttonRefr.OfficeImageId = "RefreshData";
            this.buttonRefr.ShowImage = true;
            this.buttonRefr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRefr_Click);
            // 
            // toggleButtonVerif
            // 
            this.toggleButtonVerif.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButtonVerif.Label = "Vérification de projet";
            this.toggleButtonVerif.Name = "toggleButtonVerif";
            this.toggleButtonVerif.OfficeImageId = "ReviseContents";
            this.toggleButtonVerif.ShowImage = true;
            this.toggleButtonVerif.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonVerif_Click);
            // 
            // buttonA
            // 
            this.buttonA.Label = " ";
            this.buttonA.Name = "buttonA";
            this.buttonA.OfficeImageId = "A";
            this.buttonA.ShowImage = true;
            this.buttonA.Tag = "A";
            this.buttonA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonB
            // 
            this.buttonB.Label = " ";
            this.buttonB.Name = "buttonB";
            this.buttonB.OfficeImageId = "B";
            this.buttonB.ShowImage = true;
            this.buttonB.Tag = "B";
            this.buttonB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonC
            // 
            this.buttonC.Label = " ";
            this.buttonC.Name = "buttonC";
            this.buttonC.OfficeImageId = "C";
            this.buttonC.ShowImage = true;
            this.buttonC.Tag = "C";
            this.buttonC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonD
            // 
            this.buttonD.Label = " ";
            this.buttonD.Name = "buttonD";
            this.buttonD.OfficeImageId = "D";
            this.buttonD.ShowImage = true;
            this.buttonD.Tag = "D";
            this.buttonD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonE
            // 
            this.buttonE.Label = " ";
            this.buttonE.Name = "buttonE";
            this.buttonE.OfficeImageId = "E";
            this.buttonE.ShowImage = true;
            this.buttonE.Tag = "E";
            this.buttonE.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonF
            // 
            this.buttonF.Label = " ";
            this.buttonF.Name = "buttonF";
            this.buttonF.OfficeImageId = "F";
            this.buttonF.ShowImage = true;
            this.buttonF.Tag = "F";
            this.buttonF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonG
            // 
            this.buttonG.Label = " ";
            this.buttonG.Name = "buttonG";
            this.buttonG.OfficeImageId = "G";
            this.buttonG.ShowImage = true;
            this.buttonG.Tag = "G";
            this.buttonG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonH
            // 
            this.buttonH.Label = " ";
            this.buttonH.Name = "buttonH";
            this.buttonH.OfficeImageId = "H";
            this.buttonH.ShowImage = true;
            this.buttonH.Tag = "H";
            this.buttonH.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonI
            // 
            this.buttonI.Label = " ";
            this.buttonI.Name = "buttonI";
            this.buttonI.OfficeImageId = "I";
            this.buttonI.ShowImage = true;
            this.buttonI.Tag = "I";
            this.buttonI.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonJ
            // 
            this.buttonJ.Label = " ";
            this.buttonJ.Name = "buttonJ";
            this.buttonJ.OfficeImageId = "J";
            this.buttonJ.ShowImage = true;
            this.buttonJ.Tag = "J";
            this.buttonJ.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonK
            // 
            this.buttonK.Label = " ";
            this.buttonK.Name = "buttonK";
            this.buttonK.OfficeImageId = "K";
            this.buttonK.ShowImage = true;
            this.buttonK.Tag = "K";
            this.buttonK.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonL
            // 
            this.buttonL.Label = " ";
            this.buttonL.Name = "buttonL";
            this.buttonL.OfficeImageId = "L";
            this.buttonL.ShowImage = true;
            this.buttonL.Tag = "L";
            this.buttonL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonM
            // 
            this.buttonM.Label = " ";
            this.buttonM.Name = "buttonM";
            this.buttonM.OfficeImageId = "M";
            this.buttonM.ShowImage = true;
            this.buttonM.Tag = "M";
            this.buttonM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonN
            // 
            this.buttonN.Label = " ";
            this.buttonN.Name = "buttonN";
            this.buttonN.OfficeImageId = "N";
            this.buttonN.ShowImage = true;
            this.buttonN.Tag = "N";
            this.buttonN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonO
            // 
            this.buttonO.Label = " ";
            this.buttonO.Name = "buttonO";
            this.buttonO.OfficeImageId = "O";
            this.buttonO.ShowImage = true;
            this.buttonO.Tag = "O";
            this.buttonO.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonP
            // 
            this.buttonP.Label = " ";
            this.buttonP.Name = "buttonP";
            this.buttonP.OfficeImageId = "P";
            this.buttonP.ShowImage = true;
            this.buttonP.Tag = "P";
            this.buttonP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonQ
            // 
            this.buttonQ.Label = " ";
            this.buttonQ.Name = "buttonQ";
            this.buttonQ.OfficeImageId = "Q";
            this.buttonQ.ShowImage = true;
            this.buttonQ.Tag = "Q";
            this.buttonQ.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonR
            // 
            this.buttonR.Label = " ";
            this.buttonR.Name = "buttonR";
            this.buttonR.OfficeImageId = "R";
            this.buttonR.ShowImage = true;
            this.buttonR.Tag = "R";
            this.buttonR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonS
            // 
            this.buttonS.Label = " ";
            this.buttonS.Name = "buttonS";
            this.buttonS.OfficeImageId = "S";
            this.buttonS.ShowImage = true;
            this.buttonS.Tag = "S";
            this.buttonS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonT
            // 
            this.buttonT.Label = " ";
            this.buttonT.Name = "buttonT";
            this.buttonT.OfficeImageId = "T";
            this.buttonT.ShowImage = true;
            this.buttonT.Tag = "T";
            this.buttonT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonU
            // 
            this.buttonU.Label = " ";
            this.buttonU.Name = "buttonU";
            this.buttonU.OfficeImageId = "U";
            this.buttonU.ShowImage = true;
            this.buttonU.Tag = "U";
            this.buttonU.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonV
            // 
            this.buttonV.Label = " ";
            this.buttonV.Name = "buttonV";
            this.buttonV.OfficeImageId = "V";
            this.buttonV.ShowImage = true;
            this.buttonV.Tag = "V";
            this.buttonV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonW
            // 
            this.buttonW.Label = " ";
            this.buttonW.Name = "buttonW";
            this.buttonW.OfficeImageId = "W";
            this.buttonW.ShowImage = true;
            this.buttonW.Tag = "W";
            this.buttonW.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonX
            // 
            this.buttonX.Label = " ";
            this.buttonX.Name = "buttonX";
            this.buttonX.OfficeImageId = "X";
            this.buttonX.ShowImage = true;
            this.buttonX.Tag = "X";
            this.buttonX.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonY
            // 
            this.buttonY.Label = " ";
            this.buttonY.Name = "buttonY";
            this.buttonY.OfficeImageId = "Y";
            this.buttonY.ShowImage = true;
            this.buttonY.Tag = "Y";
            this.buttonY.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonZ
            // 
            this.buttonZ.Label = " ";
            this.buttonZ.Name = "buttonZ";
            this.buttonZ.OfficeImageId = "Z";
            this.buttonZ.ShowImage = true;
            this.buttonZ.Tag = "Z";
            this.buttonZ.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // button0
            // 
            this.button0.Label = " ";
            this.button0.Name = "button0";
            this.button0.OfficeImageId = "_1";
            this.button0.ShowImage = true;
            this.button0.Tag = "1";
            this.button0.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AtoZ_Click);
            // 
            // buttonRefrRess
            // 
            this.buttonRefrRess.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRefrRess.Label = "Rafraîchir";
            this.buttonRefrRess.Name = "buttonRefrRess";
            this.buttonRefrRess.OfficeImageId = "RefreshData";
            this.buttonRefrRess.ShowImage = true;
            this.buttonRefrRess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRefr_Click);
            // 
            // buttonProdRess
            // 
            this.buttonProdRess.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonProdRess.Label = "Lister ressources";
            this.buttonProdRess.Name = "buttonProdRess";
            this.buttonProdRess.OfficeImageId = "ResourceDetailsDisplay";
            this.buttonProdRess.ShowImage = true;
            this.buttonProdRess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonProdRess_Click);
            // 
            // buttonZImpAutoRessP
            // 
            this.buttonZImpAutoRessP.Label = "Zone d\'impression automatique";
            this.buttonZImpAutoRessP.Name = "buttonZImpAutoRessP";
            this.buttonZImpAutoRessP.OfficeImageId = "ShowGridOutlook";
            this.buttonZImpAutoRessP.ShowImage = true;
            this.buttonZImpAutoRessP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonZImpAutoRessP_Click);
            // 
            // buttonIconRessP
            // 
            this.buttonIconRessP.Label = "Icône entête rapport";
            this.buttonIconRessP.Name = "buttonIconRessP";
            this.buttonIconRessP.OfficeImageId = "UpdateIcon";
            this.buttonIconRessP.ShowImage = true;
            this.buttonIconRessP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonIconRessP_Click);
            // 
            // buttonImportSoum
            // 
            this.buttonImportSoum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonImportSoum.Label = "Importation soumission";
            this.buttonImportSoum.Name = "buttonImportSoum";
            this.buttonImportSoum.OfficeImageId = "OutlineSubtotals";
            this.buttonImportSoum.ShowImage = true;
            this.buttonImportSoum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImportSoum_Click);
            // 
            // buttonImp
            // 
            this.buttonImp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonImp.Label = "Impression PDF/XLSX";
            this.buttonImp.Name = "buttonImp";
            this.buttonImp.OfficeImageId = "PrintMenu";
            this.buttonImp.ShowImage = true;
            this.buttonImp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImp_Click);
            // 
            // buttonSignAuto
            // 
            this.buttonSignAuto.Label = "Signature automatique";
            this.buttonSignAuto.Name = "buttonSignAuto";
            this.buttonSignAuto.OfficeImageId = "SignatureLineInsert";
            this.buttonSignAuto.ShowImage = true;
            this.buttonSignAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSignAuto_Click);
            // 
            // buttonZImpAuto
            // 
            this.buttonZImpAuto.Label = "Zone d\'impression automatique";
            this.buttonZImpAuto.Name = "buttonZImpAuto";
            this.buttonZImpAuto.OfficeImageId = "ShowGridOutlook";
            this.buttonZImpAuto.ShowImage = true;
            this.buttonZImpAuto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonZImpAuto_Click);
            // 
            // buttonIconB
            // 
            this.buttonIconB.Label = "Icône entête rapport";
            this.buttonIconB.Name = "buttonIconB";
            this.buttonIconB.OfficeImageId = "UpdateIcon";
            this.buttonIconB.ShowImage = true;
            this.buttonIconB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonIconB_Click);
            // 
            // buttonAffMsqUnit
            // 
            this.buttonAffMsqUnit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAffMsqUnit.Label = "Afficher/Masquer Coût unitaires";
            this.buttonAffMsqUnit.Name = "buttonAffMsqUnit";
            this.buttonAffMsqUnit.OfficeImageId = "FrameCreateLeft";
            this.buttonAffMsqUnit.ShowImage = true;
            this.buttonAffMsqUnit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAffMsqUnit_Click);
            // 
            // ManageTaskPaneRibbon
            // 
            this.Name = "ManageTaskPaneRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ManageTaskPaneRibbon_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupGestProjet.ResumeLayout(false);
            this.groupGestProjet.PerformLayout();
            this.groupDeplacement.ResumeLayout(false);
            this.groupDeplacement.PerformLayout();
            this.groupNavigation.ResumeLayout(false);
            this.groupNavigation.PerformLayout();
            this.groupSaisieItems.ResumeLayout(false);
            this.groupSaisieItems.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.groupRessources.ResumeLayout(false);
            this.groupRessources.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.buttonGroup3.ResumeLayout(false);
            this.buttonGroup3.PerformLayout();
            this.groupRessProj.ResumeLayout(false);
            this.groupRessProj.PerformLayout();
            this.groupBordereau.ResumeLayout(false);
            this.groupBordereau.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonArt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonArtAss;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAss;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLignes;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxLignes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFormula;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSupLigne;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSupProd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSupArt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAllerArt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonParam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNavFirst;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNavPrev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNavNext;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNavLast;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery galleryImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRefr;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup groupDeplacement;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup groupSaisieItems;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup groupNavigation;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonImportExcel;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonImportAccess;
        private Microsoft.Office.Tools.Ribbon.RibbonButton buttonPressePapier;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNouvP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonOuvrirP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEnregP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEnregSous;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup groupGestProjet;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonVerif;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFermerProjet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupRessources;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRefrRess;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button0;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonZ;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonY;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonX;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonW;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonV;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonU;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonQ;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonO;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonM;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonL;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonK;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonJ;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonI;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonH;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonE;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonD;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonA;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup groupRessProj;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonProdRess;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonImp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonZImpAuto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonIconB;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAffMsqUnit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSignAuto;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup groupBordereau;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonClearClip;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCut;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCopy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPaste;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonParamProjet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonImportSoum;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator9;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator10;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelBogus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonZImpAutoRessP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonIconRessP;
    }

    partial class ThisRibbonCollection
    {
        internal ManageTaskPaneRibbon ManageTaskPaneRibbon
        {
            get { return this.GetRibbon<ManageTaskPaneRibbon>(); }
        }
    }
}
