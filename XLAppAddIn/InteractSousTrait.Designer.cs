namespace XLAppAddIn
{
    partial class InteractSousTrait
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabInterSousTrait = new System.Windows.Forms.TabControl();
            this.tabConfigPart = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.tabInterSousTrait.SuspendLayout();
            this.SuspendLayout();

            // 
            // tabInterSousTrait
            // 
            this.tabInterSousTrait.Controls.Add(this.tabConfigPart);
            this.tabInterSousTrait.Controls.Add(this.tabPage2);
            this.tabInterSousTrait.Controls.Add(this.tabPage3);
            this.tabInterSousTrait.Controls.Add(this.tabPage4);
            this.tabInterSousTrait.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabInterSousTrait.Location = new System.Drawing.Point(0, 0);
            this.tabInterSousTrait.Name = "tabInterSousTrait";
            this.tabInterSousTrait.SelectedIndex = 0;
            this.tabInterSousTrait.Size = new System.Drawing.Size(1779, 658);
            this.tabInterSousTrait.TabIndex = 0;
            // 
            // tabConfigPart
            // 
            this.tabConfigPart.Location = new System.Drawing.Point(4, 22);
            this.tabConfigPart.Name = "tabConfigPart";
            this.tabConfigPart.Padding = new System.Windows.Forms.Padding(3);
            this.tabConfigPart.Size = new System.Drawing.Size(1771, 632);
            this.tabConfigPart.TabIndex = 0;
            this.tabConfigPart.Text = "Configuration partenaires";
            this.tabConfigPart.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1156, 632);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Demande de prix";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(1156, 632);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Entrée de soumission";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(1156, 632);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Analyse de soumission";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // InteractSousTrait
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabInterSousTrait);
            this.Name = "InteractSousTrait";
            this.Size = new System.Drawing.Size(1779, 658);
            this.tabInterSousTrait.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabConfigPart;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        public System.Windows.Forms.TabControl tabInterSousTrait;

    }
}
