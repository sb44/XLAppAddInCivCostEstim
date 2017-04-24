using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XLAppAddIn
{

    public partial class InterfaceVert : UserControl
    {
        const short TABLE_LAYOUT_NULL_HEIGHT = 10;
        const short TABLE_LAYOUT_FULL_HEIGHT = 312;
        private const byte PaneWidth = 135;

        public InterfaceVert()
        {
            InitializeComponent();
        }
        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        var parms = base.CreateParams;
        //        parms.Style &= ~0x02000000;  // Turn off WS_CLIPCHILDREN
        //        return parms;
        //    }
        //}
        //public Button ButtonSaisie { get { return buttonSaisie; } }
        public void EnableButtonInterfaceVert(string buttonName, string tableLayoutPanel)
        {

            var tableLayoutPanelS = activateTablePanelLayout(tableLayoutPanel) as TableLayoutPanel;
            var button = GetControlByNameInTblLayoutPan(buttonName, tableLayoutPanelS) as Button;
            button.PerformClick();
        }
        //public void EnableButton()
        //{
        //    this.buttonSaisie.Enabled = true;
        //}
        void displayColorOnClick(object sender, EventArgs e)
        {
            

            var button = sender as Button;
            if (button != null)
            {
                button.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(115)))), ((int)(((byte)(70)))));  //System.Drawing.Color.Blue;

                // mettre le control précédent de facon standard
                string buttonName = button.Name.ToString();
                adjustDisplayProperties(buttonName); 
            }     
        }
        void adjustDisplayProperties(string buttonName)
        {
            // buttonEst buttonDataProj buttonProjets buttonComm buttonDataGen buttonAdmin
            switch (buttonName)
            {
                case "buttonEst":
                    {
                        buttonDataProj.ForeColor = System.Drawing.Color.Black;
                        buttonProjets.ForeColor = System.Drawing.Color.Black;
                        buttonComm.ForeColor = System.Drawing.Color.Black;
                        //buttonDataGen.ForeColor = System.Drawing.Color.Black;
                        buttonAdmin.ForeColor = System.Drawing.Color.Black;

                        tableLayoutPanel2.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel3.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel4.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        //tableLayoutPanel5.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel6.Height = TABLE_LAYOUT_NULL_HEIGHT;

                        tableLayoutPanel2.Visible = false;
                        tableLayoutPanel3.Visible = false;
                        tableLayoutPanel4.Visible = false;
                        //tableLayoutPanel5.Visible = false;
                        tableLayoutPanel6.Visible = false;

                        tableLayoutPanel1.Height = TABLE_LAYOUT_FULL_HEIGHT;
                        tableLayoutPanel1.Visible = true;

                        break;
                    }
                case "buttonDataProj":
                    {
                        // sous boutons - reset style
                        buttonSrcAss.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(243)))), ((int)(((byte)(243)))));
                        buttonSrcRess.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(243)))), ((int)(((byte)(243)))));
                        buttonSrcAss.FlatAppearance.BorderSize = 0;
                        buttonSrcRess.FlatAppearance.BorderSize = 0;
                        // fin sous boutons

                        buttonEst.ForeColor = System.Drawing.Color.Black;
                        buttonProjets.ForeColor = System.Drawing.Color.Black;
                        buttonComm.ForeColor = System.Drawing.Color.Black;
                        //buttonDataGen.ForeColor = System.Drawing.Color.Black;
                        buttonAdmin.ForeColor = System.Drawing.Color.Black;

                        tableLayoutPanel1.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel3.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel4.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        //tableLayoutPanel5.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel6.Height = TABLE_LAYOUT_NULL_HEIGHT;

                        tableLayoutPanel1.Visible = false;
                        tableLayoutPanel3.Visible = false;
                        tableLayoutPanel4.Visible = false;
                        //tableLayoutPanel5.Visible = false;
                        tableLayoutPanel6.Visible = false;

                        tableLayoutPanel2.Height = TABLE_LAYOUT_FULL_HEIGHT;
                        tableLayoutPanel2.Visible = true;

                        break;
                    }
                case "buttonProjets":
                    {
                        buttonDataProj.ForeColor = System.Drawing.Color.Black;
                        buttonEst.ForeColor = System.Drawing.Color.Black;
                        buttonComm.ForeColor = System.Drawing.Color.Black;
                        //buttonDataGen.ForeColor = System.Drawing.Color.Black;
                        buttonAdmin.ForeColor = System.Drawing.Color.Black;

                        tableLayoutPanel1.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel2.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel4.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        //tableLayoutPanel5.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel6.Height = TABLE_LAYOUT_NULL_HEIGHT;

                        tableLayoutPanel1.Visible = false;
                        tableLayoutPanel2.Visible = false;
                        tableLayoutPanel4.Visible = false;
                        //tableLayoutPanel5.Visible = false;
                        tableLayoutPanel6.Visible = false;

                        tableLayoutPanel3.Height = TABLE_LAYOUT_FULL_HEIGHT;
                        tableLayoutPanel3.Visible = true;

                        break;
                    }
                case "buttonComm":
                    {
                        buttonDataProj.ForeColor = System.Drawing.Color.Black;
                        buttonProjets.ForeColor = System.Drawing.Color.Black;
                        buttonEst.ForeColor = System.Drawing.Color.Black;
                        //buttonDataGen.ForeColor = System.Drawing.Color.Black;
                        buttonAdmin.ForeColor = System.Drawing.Color.Black;

                        tableLayoutPanel1.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel2.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel3.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        //tableLayoutPanel5.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel6.Height = TABLE_LAYOUT_NULL_HEIGHT;

                        tableLayoutPanel1.Visible = false;
                        tableLayoutPanel2.Visible = false;
                        tableLayoutPanel3.Visible = false;
                        //tableLayoutPanel5.Visible = false;
                        tableLayoutPanel6.Visible = false;

                        tableLayoutPanel4.Height = TABLE_LAYOUT_FULL_HEIGHT;
                        tableLayoutPanel4.Visible = true;

                        break;
                    }
                //case "buttonDataGen":
                //    {
                //        buttonDataProj.ForeColor = System.Drawing.Color.Black;
                //        buttonProjets.ForeColor = System.Drawing.Color.Black;
                //        buttonComm.ForeColor = System.Drawing.Color.Black;
                //        buttonEst.ForeColor = System.Drawing.Color.Black;
                //        buttonAdmin.ForeColor = System.Drawing.Color.Black;

                //        tableLayoutPanel1.Height = TABLE_LAYOUT_NULL_HEIGHT;
                //        tableLayoutPanel2.Height = TABLE_LAYOUT_NULL_HEIGHT;
                //        tableLayoutPanel3.Height = TABLE_LAYOUT_NULL_HEIGHT;
                //        tableLayoutPanel4.Height = TABLE_LAYOUT_NULL_HEIGHT;
                //        tableLayoutPanel6.Height = TABLE_LAYOUT_NULL_HEIGHT;

                //        tableLayoutPanel1.Visible = false;
                //        tableLayoutPanel2.Visible = false;
                //        tableLayoutPanel3.Visible = false;
                //        tableLayoutPanel4.Visible = false;
                //        tableLayoutPanel6.Visible = false;

                //        //tableLayoutPanel5.Height = TABLE_LAYOUT_FULL_HEIGHT;
                //        //tableLayoutPanel5.Visible = true;

                //        break;
                //    }
                case "buttonAdmin":
                    {
                        // sous boutons - reset style :
                        buttonMateriaux.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(243)))), ((int)(((byte)(243)))));
                        buttonBordereau.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(243)))), ((int)(((byte)(243)))));
                        buttonMateriaux.FlatAppearance.BorderSize = 0;
                        buttonBordereau.FlatAppearance.BorderSize = 0;
                        // fin sous boutons

                        buttonDataProj.ForeColor = System.Drawing.Color.Black;
                        buttonProjets.ForeColor = System.Drawing.Color.Black;
                        buttonComm.ForeColor = System.Drawing.Color.Black;
                        //buttonDataGen.ForeColor = System.Drawing.Color.Black;
                        buttonEst.ForeColor = System.Drawing.Color.Black;

                        tableLayoutPanel1.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel2.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel3.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        tableLayoutPanel4.Height = TABLE_LAYOUT_NULL_HEIGHT;
                        //tableLayoutPanel5.Height = TABLE_LAYOUT_NULL_HEIGHT;

                        tableLayoutPanel1.Visible = false;
                        tableLayoutPanel2.Visible = false;
                        tableLayoutPanel3.Visible = false;
                        tableLayoutPanel4.Visible = false;
                        //tableLayoutPanel5.Visible = false;

                        tableLayoutPanel6.Height = TABLE_LAYOUT_FULL_HEIGHT;
                        tableLayoutPanel6.Visible = true;

                        break;
                    }
                default:
                    break;
            }
        }
        /*
        void ShowSelectionPictureBox(object sender, EventArgs e)
        {
            //-2; 96 -- pictureBoxShowSelection
                        var button = sender as Button;
                        if (button != null)
                        {


                            pictureBoxShowSelection.Location = new Point(-3, (button.Top + button.Height)/2);
                            pictureBoxShowSelection.Width = 12;
                            pictureBoxShowSelection.Height = 12;

                            pictureBoxShowSelection.Visible = true;
                            pictureBoxShowSelection.Invalidate();

                        }
            
        }
        */
        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void MyUserControl_Load(object sender, EventArgs e)
        {
            this.SizeChanged += new EventHandler(MyUserControl_SizeChanged);
            buttonEst_Click(sender, e);

            //À TESTER - TODO : VISIBLE EVENTHANDLER FOR GROUP TABS ON RIBBON

            //buttonSaisie_Click(sender, e); enlevé le 26-11-2016 testing

            //this.Dock = DockStyle.Bottom;
            // this.Dock = DockStyle.Fill;
        }

        void MyUserControl_SizeChanged(object sender, EventArgs e)
            {
            
            if (Globals.ThisAddIn.myCustomTaskPane != null && Globals.ThisAddIn.myCustomTaskPane.Visible && Globals.ThisAddIn.myCustomTaskPane.Width != PaneWidth)
            {
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;
                System.Windows.Forms.SendKeys.Send("{ESC}");
                //Globals.ThisAddIn.myCustomTaskPane.Visible = false;
                Globals.ThisAddIn.myCustomTaskPane.Width = PaneWidth;
                //Globals.ThisAddIn.myCustomTaskPane.Visible = true;
                System.Windows.Forms.SendKeys.Send("{ESC}");
                Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
            }
            
        }

        public void buttonEst_Click(object sender, EventArgs e)
        {
            displayColorOnClick(buttonEst, e);

        }

        private void buttonDataProj_Click(object sender, EventArgs e)
        {
            displayColorOnClick(buttonDataProj, e);
        }

        private void buttonProjets_Click(object sender, EventArgs e)
        {
            displayColorOnClick(buttonProjets, e);
        }

        private void buttonComm_Click(object sender, EventArgs e)
        {
            displayColorOnClick(buttonComm, e);
        }

        //private void buttonDataGen_Click(object sender, EventArgs e)
        //{
        //    displayColorOnClick(buttonDataGen, e);
        //}

        private void buttonAdmin_Click(object sender, EventArgs e)
        {
            displayColorOnClick(buttonAdmin, e);

        }

        private void buttonNextGroup_Click(object sender, EventArgs e)
        {
        // buttonEst buttonDataProj buttonProjets buttonComm buttonDataGen buttonAdmin

            //var button = sender as Button;
        if (tableLayoutPanel1.Height != TABLE_LAYOUT_NULL_HEIGHT)
            buttonDataProj_Click(buttonDataProj, e);
        else if (tableLayoutPanel2.Height != TABLE_LAYOUT_NULL_HEIGHT)
            buttonProjets_Click(buttonProjets, e);
        else if (tableLayoutPanel3.Height != TABLE_LAYOUT_NULL_HEIGHT)
            buttonComm_Click(buttonComm, e);
        else if (tableLayoutPanel4.Height != TABLE_LAYOUT_NULL_HEIGHT)
                buttonAdmin_Click(buttonAdmin, e);
        //else if (tableLayoutPanel5.Height != TABLE_LAYOUT_NULL_HEIGHT)
        //    buttonAdmin_Click(buttonAdmin, e);
            else if (tableLayoutPanel6.Height != TABLE_LAYOUT_NULL_HEIGHT)
            buttonEst_Click(buttonEst, e);
        }

        private void buttonPreviousGroup_Click(object sender, EventArgs e)
        {
        // 1 buttonEst 2 buttonDataProj 3 buttonProjets 4 buttonComm 5 buttonDataGen 6 buttonAdmin

           // var button = sender as Button;
            if (tableLayoutPanel1.Height != TABLE_LAYOUT_NULL_HEIGHT)
                buttonAdmin_Click(buttonAdmin, e);
            else if (tableLayoutPanel2.Height != TABLE_LAYOUT_NULL_HEIGHT)
                buttonEst_Click(buttonEst, e);
            else if (tableLayoutPanel3.Height != TABLE_LAYOUT_NULL_HEIGHT)
                buttonDataProj_Click(buttonDataProj, e);
            else if (tableLayoutPanel4.Height != TABLE_LAYOUT_NULL_HEIGHT)
                buttonProjets_Click(buttonProjets, e);
            //else if (tableLayoutPanel5.Height != TABLE_LAYOUT_NULL_HEIGHT)
            //    buttonComm_Click(buttonComm, e);
            else if (tableLayoutPanel6.Height != TABLE_LAYOUT_NULL_HEIGHT)
                buttonComm_Click(buttonComm, e);
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            if (tableLayoutPanel1.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel1, e);
            else if (tableLayoutPanel2.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel2, e);
            else if (tableLayoutPanel3.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel3, e);
            else if (tableLayoutPanel4.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel4, e);
            //else if (tableLayoutPanel5.Height != TABLE_LAYOUT_NULL_HEIGHT)
            //    toggleButtonSelection(tableLayoutPanel5, e);
            else if (tableLayoutPanel6.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel6, e);
        }
        private TableLayoutPanel determineActiveTablePanelLayout()
        {
            if (tableLayoutPanel1.Height != TABLE_LAYOUT_NULL_HEIGHT)
                return tableLayoutPanel1;
            else if (tableLayoutPanel2.Height != TABLE_LAYOUT_NULL_HEIGHT)
                return tableLayoutPanel2;
            else if (tableLayoutPanel3.Height != TABLE_LAYOUT_NULL_HEIGHT)
                return tableLayoutPanel3;
            else if (tableLayoutPanel4.Height != TABLE_LAYOUT_NULL_HEIGHT)
                return tableLayoutPanel4;
            //else if (tableLayoutPanel5.Height != TABLE_LAYOUT_NULL_HEIGHT)
            //    return tableLayoutPanel5;
            else if (tableLayoutPanel6.Height != TABLE_LAYOUT_NULL_HEIGHT)
                return tableLayoutPanel6;
            else
                return tableLayoutPanel1;
        }

        private TableLayoutPanel activateTablePanelLayout(string tableLayoutPanelstr)
        {
            if (tableLayoutPanelstr == "tableLayoutPanel1")
                return tableLayoutPanel1;
            else if (tableLayoutPanelstr == "tableLayoutPanel2")
                return tableLayoutPanel2;
            else if (tableLayoutPanelstr == "tableLayoutPanel3")
                return tableLayoutPanel3;
            else if (tableLayoutPanelstr == "tableLayoutPanel4")
                return tableLayoutPanel4;
            //else if (tableLayoutPanelstr == "tableLayoutPanel5")
            //    return tableLayoutPanel5;
            else if (tableLayoutPanelstr == "tableLayoutPanel6")
                return tableLayoutPanel6;
            else
                return tableLayoutPanel1;
        }

        Control GetControlByName(string Name)
        {
            foreach (Control c in this.Controls)
                if (c.Name == Name)
                    return c;

            return null;
        }

        Control GetControlByNameInTblLayoutPan(string Name, TableLayoutPanel tableLayoutPanel)
        {
            foreach (Button b in tableLayoutPanel.Controls)
                if (b.Name == Name)
                    return b;

            return null;
        }
        Control GetControlByTagInTblLayoutPan(string buttonTag, TableLayoutPanel tableLayoutPanel, bool MoveDown = true)
        {
            // lastItemInTab  beforeLastItemInTab  firstItemInTab
            if (buttonTag == "beforeLastItemInTab" && MoveDown)
            {
                foreach (Button b in tableLayoutPanel.Controls)
                    if (b.Tag.ToString() == "lastItemInTab")
                        return b;
            }
            else if (buttonTag == "lastItemInTab" && MoveDown)
            {
               foreach (Button b in tableLayoutPanel.Controls)
                   if (b.Tag.ToString() == "firstItemInTab")
                        return b;
            }
            else if (buttonTag == "lastItemInTab" && !MoveDown)
            {
                foreach (Button b in tableLayoutPanel.Controls)
                    if (b.Tag.ToString() == "beforeLastItemInTab")
                        return b;
            }
            return null;
        }

        string GetLastButtonName(string buttonTag, TableLayoutPanel tableLayoutPanel, bool MoveDown = true)
        {
            foreach (Button b in tableLayoutPanel.Controls)
                if (b.Tag.ToString() == "lastItemInTab")
                    return b.Name.ToString();
            return null;
        }

        void toggleButtonSelection(object sender, EventArgs e, bool MoveDown = true)
        {
            var tableLayoutPanel = sender as TableLayoutPanel;
            if (tableLayoutPanel != null)
            {

                bool notFound = true;
                string buttonPrevName = "";

                foreach (Button button in tableLayoutPanel.Controls) //LOOPING ds la tableLayoutPanel active
                {
                    
                    if (!notFound && MoveDown)
                    {
                        button.PerformClick();
                        break;
                    }
                    //if (button.ForeColor == System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(115)))), ((int)(((byte)(70)))))) notFound = false;
                    if (button.FlatAppearance.BorderSize == 1) notFound = false;
                    if (!notFound && !MoveDown)
                    {
                        // go back one...
                        if (buttonPrevName == "")
                            buttonPrevName = GetLastButtonName(button.Tag.ToString(), tableLayoutPanel, MoveDown);  // if first.. get By tag - go to Last
                          
                        var prevButton = GetControlByNameInTblLayoutPan(buttonPrevName, tableLayoutPanel) as Button;
                        
                        prevButton.PerformClick();
                        break;
                    }
                    buttonPrevName = button.Name.ToString();
                    // BIEN identifié les tabs...
                    // lastItemInTab firstItemInTab beforeLastItemInTab
                    //if (button.Name == "buttonFerm" && MoveDown) buttonSaisie_Click(buttonSaisie, e); // À FAIRE : identifier si extremum et clicker en conséquence du sens choisi selon la tableLayoutPanel active
                    //if (button.Name == "buttonFerm" && !MoveDown) buttonDemEnv_Click(buttonDemEnv, e); // À FAIRE : identifier si extremum et clicker en conséquence du sens choisi selon la tableLayoutPanel active
                    if ((button.Tag.ToString() == "lastItemInTab" && MoveDown) || (button.Tag.ToString() == "lastItemInTab" && !MoveDown))
                    {
                        var toSelectButton = GetControlByTagInTblLayoutPan(button.Tag.ToString(), tableLayoutPanel, MoveDown) as Button;
                        toSelectButton.PerformClick();
                    } 

                }
                /*
                if (notFound && MoveDown) 
                    buttonSaisie_Click(buttonSaisie, e);
                else if (notFound && !MoveDown)
                    buttonDemEnv_Click(buttonDemEnv, e);
                */
            }
            else // StandardOnClick... :
            {
                var tableLayoutPanelActive = determineActiveTablePanelLayout() as TableLayoutPanel;
                var selectedButton = sender as Button;
                if (selectedButton != null)
                {
                        foreach (Button button in tableLayoutPanelActive.Controls)
                        {
                        //if (button.ForeColor == System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(115)))), ((int)(((byte)(70))))))
                        if (button.FlatAppearance.BorderSize == 1)
                        {
                            // RESET style
                                button.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(243)))), ((int)(((byte)(243)))));
                            //    button.ForeColor = System.Drawing.Color.Black;
                                button.FlatAppearance.BorderSize = 0;
                                break;
                            }

                        }
                        // SET style
                        selectedButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(159)))), ((int)(((byte)(213)))), ((int)(((byte)(183)))));
                       // selectedButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(115)))), ((int)(((byte)(70)))));
                        selectedButton.FlatAppearance.BorderSize = 1;
                }
                
            }
        }

        private void buttonPrevious_Click(object sender, EventArgs e)
        {
            bool MoveDown = false;
            if (tableLayoutPanel1.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel1, e, MoveDown);
            else if (tableLayoutPanel2.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel2, e, MoveDown);
            else if (tableLayoutPanel3.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel3, e, MoveDown);
            else if (tableLayoutPanel4.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel4, e, MoveDown);
            //else if (tableLayoutPanel5.Height != TABLE_LAYOUT_NULL_HEIGHT)
            //    toggleButtonSelection(tableLayoutPanel5, e, MoveDown);
            else if (tableLayoutPanel6.Height != TABLE_LAYOUT_NULL_HEIGHT)
                toggleButtonSelection(tableLayoutPanel6, e, MoveDown);
        }
        void toggleVisibileRibbonGroupTabs(object sender, EventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                // mettre le control précédent de facon standard
                string userControlButtonName = button.Name.ToString();
                adjustRibbonGroupTabVisibleProperty(userControlButtonName);
            }
        }
        static void adjustRibbonGroupTabVisibleProperty(string userControlButtonName = "")
        {
            // buttonEst buttonDataProj buttonProjets buttonComm buttonDataGen buttonAdmin

            switch (userControlButtonName)
            {
                case "buttonSaisie":
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessources.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = true;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = true;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = true;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = true;
                        break;
                    }
                case "buttonRess":
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessources.Visible = true;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = true;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = true;
                        break;
                    }
                case "buttonExc":
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessources.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = true;
                        break;
                    }
                case "buttonDemEnv":
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessources.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = true;
                        break;
                    }
                case "buttonFerm":
                    {
                        //Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = true;
                        //Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = true;
                        //Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = true;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = true;
                        break;
                    }
                case "buttonGestFichProj":
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = true;
                        break;
                    
                    }
                case "buttonMateriaux":
                    {
                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = true;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = true;

                        break;
                    }
                case "buttonBordereau":
                    {
                        //TODO : LES COMMANDES du rubban pour cette feuille!

                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = true;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = true;
                        break;
                    }
                default:
                    {
                        //all invisible except App XL office 365
                        Globals.Ribbons.ManageTaskPaneRibbon.groupRessProj.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupNavigation.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupSaisieItems.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupGestProjet.Visible = false;
                        Globals.Ribbons.ManageTaskPaneRibbon.groupBordereau.Visible = false;

                        Globals.Ribbons.ManageTaskPaneRibbon.groupDeplacement.Visible = true;

                        break;
                    }
            }
        }

        public void buttonSaisie_Click(object sender, EventArgs e)
        {
            
            toggleButtonSelection(buttonSaisie, e);
            toggleVisibileRibbonGroupTabs(buttonSaisie, e);
            Globals.ThisAddIn.Application.Run("ShowSaisieSheet");
        }

        private void buttonRess_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonRess, e);
            toggleVisibileRibbonGroupTabs(buttonRess, e);
            Globals.ThisAddIn.Application.Run("ShowProduitsSheet");
        }

        private void buttonExc_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonExc, e);
            toggleVisibileRibbonGroupTabs(buttonExc, e);
            Globals.ThisAddIn.Application.Run("ShowTrancheSheet");
        }

        private void buttonDemEnv_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonDemEnv, e);
            toggleVisibileRibbonGroupTabs(buttonDemEnv, e);
            
            //method à développer pour gérer la visibilité
            Globals.ThisAddIn.TaskPaneInterfaceData.Visible = true;
            Globals.ThisAddIn.TaskPaneInterfaceSousTrait.Visible = true;
           //Globals.ThisAddIn.TaskPaneVerifProjet.Visible = true;
            
            
            
        }

        private void buttonFerm_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonFerm, e);
            //toggleVisibileRibbonGroupTabs(buttonFerm, e); //mm que saisie...
            
            //System.Threading.Thread.Sleep(50);
            Application.DoEvents(); //update le GUID avant de caller la macroVBA
            
            Globals.ThisAddIn.Application.Run("FermetureSoum");
        }

        private void buttonSrcRess_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonSrcRess, e);
            Globals.ThisAddIn.Application.Run("configRess");
            buttonDataProj.PerformClick();

        }

        private void buttonbuttonSrcAss_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonSrcAss, e);
            Globals.ThisAddIn.Application.Run("configBD");
            buttonDataProj.PerformClick();
        }

        private void buttonGestFichProj_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonGestFichProj, e);
            toggleVisibileRibbonGroupTabs(buttonGestFichProj, e);
        }

        private void buttonBordereau_Click(object sender, EventArgs e)
        {
            
            toggleButtonSelection(buttonBordereau, e);
            //TODO : LES COMMANDES du rubban pour cette feuille!
            toggleVisibileRibbonGroupTabs(buttonBordereau, e);

            Globals.ThisAddIn.Application.Run("ShowBorderauSheet");
        }

        private void buttonMateriaux_Click(object sender, EventArgs e)
        {
            toggleButtonSelection(buttonMateriaux, e);
            toggleVisibileRibbonGroupTabs(buttonMateriaux, e);
            Globals.ThisAddIn.Application.Run("ShowRessourcesSheet");
            //RAPPORTRESSOURCES
        }
    }
}
