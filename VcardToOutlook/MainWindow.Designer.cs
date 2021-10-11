namespace VcardToOutlook
{
    partial class MainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxOutput = new System.Windows.Forms.TextBox();
            this.textBoxInput = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.buttonSelectTarget = new System.Windows.Forms.Button();
            this.buttonSelectSource = new System.Windows.Forms.Button();
            this.buttonImport = new System.Windows.Forms.Button();
            this.buttonAbout = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.buttonCut = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.checkBoxClearOldVcf = new System.Windows.Forms.CheckBox();
            this.checkBoxClearOldContact = new System.Windows.Forms.CheckBox();
            this.linkLabelWebsite = new System.Windows.Forms.LinkLabel();
            this.label6 = new System.Windows.Forms.Label();
            this.checkBoxRemoveVietnameseSign = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(221, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(171, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "VCardToOutlook 1.1";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(186, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(225, 47);
            this.label2.TabIndex = 3;
            this.label2.Text = "Select a VCF file and then click on \'Cut\' to split it into individual VCF files. " +
    "Then you can copy them to your phone.";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(183, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(222, 32);
            this.label4.TabIndex = 5;
            this.label4.Text = "You can also import to outlook with one-click.";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBoxOutput);
            this.groupBox1.Controls.Add(this.textBoxInput);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.buttonSelectTarget);
            this.groupBox1.Controls.Add(this.buttonSelectSource);
            this.groupBox1.Location = new System.Drawing.Point(22, 142);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(389, 116);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Provide details";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 56);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Output Folder";
            // 
            // textBoxOutput
            // 
            this.textBoxOutput.Location = new System.Drawing.Point(19, 81);
            this.textBoxOutput.Name = "textBoxOutput";
            this.textBoxOutput.Size = new System.Drawing.Size(334, 20);
            this.textBoxOutput.TabIndex = 5;
            // 
            // textBoxInput
            // 
            this.textBoxInput.Location = new System.Drawing.Point(19, 33);
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(334, 20);
            this.textBoxInput.TabIndex = 3;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 16);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(135, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Source (Original vCard File)";
            // 
            // buttonSelectTarget
            // 
            this.buttonSelectTarget.Image = global::VcardToOutlook.Properties.Resources.output;
            this.buttonSelectTarget.Location = new System.Drawing.Point(359, 78);
            this.buttonSelectTarget.Name = "buttonSelectTarget";
            this.buttonSelectTarget.Size = new System.Drawing.Size(24, 23);
            this.buttonSelectTarget.TabIndex = 4;
            this.buttonSelectTarget.UseVisualStyleBackColor = true;
            this.buttonSelectTarget.Click += new System.EventHandler(this.buttonSelectTarget_Click);
            // 
            // buttonSelectSource
            // 
            this.buttonSelectSource.Image = global::VcardToOutlook.Properties.Resources.open;
            this.buttonSelectSource.Location = new System.Drawing.Point(359, 30);
            this.buttonSelectSource.Name = "buttonSelectSource";
            this.buttonSelectSource.Size = new System.Drawing.Size(24, 23);
            this.buttonSelectSource.TabIndex = 0;
            this.buttonSelectSource.UseVisualStyleBackColor = true;
            this.buttonSelectSource.Click += new System.EventHandler(this.buttonSelectSource_Click);
            // 
            // buttonImport
            // 
            this.buttonImport.Location = new System.Drawing.Point(137, 308);
            this.buttonImport.Name = "buttonImport";
            this.buttonImport.Size = new System.Drawing.Size(70, 22);
            this.buttonImport.TabIndex = 8;
            this.buttonImport.Text = "Import";
            this.buttonImport.UseVisualStyleBackColor = true;
            this.buttonImport.Click += new System.EventHandler(this.buttonImport_Click);
            // 
            // buttonAbout
            // 
            this.buttonAbout.Location = new System.Drawing.Point(341, 114);
            this.buttonAbout.Name = "buttonAbout";
            this.buttonAbout.Size = new System.Drawing.Size(70, 22);
            this.buttonAbout.TabIndex = 9;
            this.buttonAbout.Text = "About";
            this.buttonAbout.UseVisualStyleBackColor = true;
            this.buttonAbout.Click += new System.EventHandler(this.buttonAbout_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(213, 307);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(196, 22);
            this.progressBar.TabIndex = 10;
            this.progressBar.Visible = false;
            // 
            // buttonCut
            // 
            this.buttonCut.Location = new System.Drawing.Point(137, 284);
            this.buttonCut.Name = "buttonCut";
            this.buttonCut.Size = new System.Drawing.Size(70, 22);
            this.buttonCut.TabIndex = 7;
            this.buttonCut.Text = "Cut";
            this.buttonCut.UseVisualStyleBackColor = true;
            this.buttonCut.Click += new System.EventHandler(this.buttonCut_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::VcardToOutlook.Properties.Resources.Microsoft_outlook;
            this.pictureBox1.Location = new System.Drawing.Point(22, 22);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(158, 103);
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // checkBoxClearOldVcf
            // 
            this.checkBoxClearOldVcf.AutoSize = true;
            this.checkBoxClearOldVcf.Location = new System.Drawing.Point(20, 289);
            this.checkBoxClearOldVcf.Name = "checkBoxClearOldVcf";
            this.checkBoxClearOldVcf.Size = new System.Drawing.Size(112, 17);
            this.checkBoxClearOldVcf.TabIndex = 11;
            this.checkBoxClearOldVcf.Text = "Clear Old Vcf Files";
            this.checkBoxClearOldVcf.UseVisualStyleBackColor = true;
            // 
            // checkBoxClearOldContact
            // 
            this.checkBoxClearOldContact.AutoSize = true;
            this.checkBoxClearOldContact.Location = new System.Drawing.Point(20, 312);
            this.checkBoxClearOldContact.Name = "checkBoxClearOldContact";
            this.checkBoxClearOldContact.Size = new System.Drawing.Size(109, 17);
            this.checkBoxClearOldContact.TabIndex = 12;
            this.checkBoxClearOldContact.Text = "Clear Old Contact";
            this.checkBoxClearOldContact.UseVisualStyleBackColor = true;
            // 
            // linkLabelWebsite
            // 
            this.linkLabelWebsite.AutoSize = true;
            this.linkLabelWebsite.Location = new System.Drawing.Point(248, 122);
            this.linkLabelWebsite.Name = "linkLabelWebsite";
            this.linkLabelWebsite.Size = new System.Drawing.Size(90, 13);
            this.linkLabelWebsite.TabIndex = 13;
            this.linkLabelWebsite.TabStop = true;
            this.linkLabelWebsite.Text = "https://bbhcm.vn";
            this.linkLabelWebsite.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelWebsite_LinkClicked);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(212, 291);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "label status";
            // 
            // checkBoxRemoveVietnameseSign
            // 
            this.checkBoxRemoveVietnameseSign.AutoSize = true;
            this.checkBoxRemoveVietnameseSign.Location = new System.Drawing.Point(20, 264);
            this.checkBoxRemoveVietnameseSign.Name = "checkBoxRemoveVietnameseSign";
            this.checkBoxRemoveVietnameseSign.Size = new System.Drawing.Size(146, 17);
            this.checkBoxRemoveVietnameseSign.TabIndex = 15;
            this.checkBoxRemoveVietnameseSign.Text = "Remove Vietnamese sign";
            this.checkBoxRemoveVietnameseSign.UseVisualStyleBackColor = true;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(423, 341);
            this.Controls.Add(this.checkBoxRemoveVietnameseSign);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.linkLabelWebsite);
            this.Controls.Add(this.checkBoxClearOldContact);
            this.Controls.Add(this.checkBoxClearOldVcf);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.buttonAbout);
            this.Controls.Add(this.buttonImport);
            this.Controls.Add(this.buttonCut);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VCard To OutLook";
            this.Load += new System.EventHandler(this.MainWindow_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonSelectSource;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxInput;
        private System.Windows.Forms.TextBox textBoxOutput;
        private System.Windows.Forms.Button buttonSelectTarget;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonCut;
        private System.Windows.Forms.Button buttonImport;
        private System.Windows.Forms.Button buttonAbout;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.CheckBox checkBoxClearOldVcf;
        private System.Windows.Forms.CheckBox checkBoxClearOldContact;
        private System.Windows.Forms.LinkLabel linkLabelWebsite;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox checkBoxRemoveVietnameseSign;
    }
}

