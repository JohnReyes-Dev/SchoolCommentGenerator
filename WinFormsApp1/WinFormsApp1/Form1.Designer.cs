namespace WinFormsApp1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.NextStudentButton = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.PreviousStudentButton = new System.Windows.Forms.Button();
            this.CreateFilesForEvaluationButton = new System.Windows.Forms.Button();
            this.GetExelFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.OutHowManyStudents = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.FolderLocationName = new System.Windows.Forms.Label();
            this.WhereToPutTheFilesButton = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.EvaluationFileTemplateName = new System.Windows.Forms.Label();
            this.FindEvaluationFileTemplateButton = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.ErrorMessageLabel = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.richTextBox1.Location = new System.Drawing.Point(68, 73);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(335, 177);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // NextStudentButton
            // 
            this.NextStudentButton.Enabled = false;
            this.NextStudentButton.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.NextStudentButton.Location = new System.Drawing.Point(325, 17);
            this.NextStudentButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.NextStudentButton.Name = "NextStudentButton";
            this.NextStudentButton.Size = new System.Drawing.Size(78, 47);
            this.NextStudentButton.TabIndex = 2;
            this.NextStudentButton.Text = "---->";
            this.NextStudentButton.UseVisualStyleBackColor = true;
            this.NextStudentButton.Click += new System.EventHandler(this.NextStudentButton_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.Enabled = false;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.IntegralHeight = false;
            this.comboBox1.Location = new System.Drawing.Point(444, 87);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(151, 23);
            this.comboBox1.TabIndex = 3;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // PreviousStudentButton
            // 
            this.PreviousStudentButton.Enabled = false;
            this.PreviousStudentButton.Font = new System.Drawing.Font("Segoe UI", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.PreviousStudentButton.Location = new System.Drawing.Point(224, 17);
            this.PreviousStudentButton.Name = "PreviousStudentButton";
            this.PreviousStudentButton.Size = new System.Drawing.Size(80, 47);
            this.PreviousStudentButton.TabIndex = 4;
            this.PreviousStudentButton.Text = "<----";
            this.PreviousStudentButton.UseVisualStyleBackColor = true;
            this.PreviousStudentButton.Click += new System.EventHandler(this.PreviousStudent_Click);
            // 
            // CreateFilesForEvaluationButton
            // 
            this.CreateFilesForEvaluationButton.Enabled = false;
            this.CreateFilesForEvaluationButton.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.CreateFilesForEvaluationButton.Location = new System.Drawing.Point(120, 28);
            this.CreateFilesForEvaluationButton.Name = "CreateFilesForEvaluationButton";
            this.CreateFilesForEvaluationButton.Size = new System.Drawing.Size(116, 68);
            this.CreateFilesForEvaluationButton.TabIndex = 5;
            this.CreateFilesForEvaluationButton.Text = "Create files";
            this.CreateFilesForEvaluationButton.UseVisualStyleBackColor = true;
            this.CreateFilesForEvaluationButton.Click += new System.EventHandler(this.CreateFilesForEvaluationButton_Click);
            // 
            // GetExelFile
            // 
            this.GetExelFile.Location = new System.Drawing.Point(444, 22);
            this.GetExelFile.Name = "GetExelFile";
            this.GetExelFile.Size = new System.Drawing.Size(151, 49);
            this.GetExelFile.TabIndex = 6;
            this.GetExelFile.Text = "GetFile";
            this.GetExelFile.UseVisualStyleBackColor = true;
            this.GetExelFile.Click += new System.EventHandler(this.GetExelFileButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(336, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 32);
            this.label1.TabIndex = 7;
            this.label1.Text = "Step 1:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(340, 84);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 32);
            this.label2.TabIndex = 8;
            this.label2.Text = "Step 2:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(26, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 32);
            this.label3.TabIndex = 9;
            // 
            // OutHowManyStudents
            // 
            this.OutHowManyStudents.AutoSize = true;
            this.OutHowManyStudents.Font = new System.Drawing.Font("Segoe UI", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.OutHowManyStudents.Location = new System.Drawing.Point(325, 252);
            this.OutHowManyStudents.Name = "OutHowManyStudents";
            this.OutHowManyStudents.Size = new System.Drawing.Size(0, 25);
            this.OutHowManyStudents.TabIndex = 10;
            this.OutHowManyStudents.Click += new System.EventHandler(this.OutHowManyStudents_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.groupBox1.Controls.Add(this.richTextBox1);
            this.groupBox1.Controls.Add(this.NextStudentButton);
            this.groupBox1.Controls.Add(this.OutHowManyStudents);
            this.groupBox1.Controls.Add(this.PreviousStudentButton);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBox1.Location = new System.Drawing.Point(7, 171);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(458, 291);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Step 3";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.groupBox6);
            this.groupBox2.Controls.Add(this.groupBox5);
            this.groupBox2.Controls.Add(this.groupBox4);
            this.groupBox2.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBox2.Location = new System.Drawing.Point(492, 171);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(388, 459);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Step 3";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.CreateFilesForEvaluationButton);
            this.groupBox6.Location = new System.Drawing.Point(6, 339);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(376, 114);
            this.groupBox6.TabIndex = 15;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Step 3";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.FolderLocationName);
            this.groupBox5.Controls.Add(this.WhereToPutTheFilesButton);
            this.groupBox5.Location = new System.Drawing.Point(6, 188);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(376, 139);
            this.groupBox5.TabIndex = 15;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Step 2";
            // 
            // FolderLocationName
            // 
            this.FolderLocationName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.FolderLocationName.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.FolderLocationName.Location = new System.Drawing.Point(6, 103);
            this.FolderLocationName.Name = "FolderLocationName";
            this.FolderLocationName.Size = new System.Drawing.Size(364, 24);
            this.FolderLocationName.TabIndex = 20;
            this.FolderLocationName.Text = "Folder Location Name?";
            // 
            // WhereToPutTheFilesButton
            // 
            this.WhereToPutTheFilesButton.Enabled = false;
            this.WhereToPutTheFilesButton.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.WhereToPutTheFilesButton.Location = new System.Drawing.Point(6, 38);
            this.WhereToPutTheFilesButton.Name = "WhereToPutTheFilesButton";
            this.WhereToPutTheFilesButton.Size = new System.Drawing.Size(170, 49);
            this.WhereToPutTheFilesButton.TabIndex = 6;
            this.WhereToPutTheFilesButton.Text = "Where to Put the Files?";
            this.WhereToPutTheFilesButton.UseVisualStyleBackColor = true;
            this.WhereToPutTheFilesButton.Click += new System.EventHandler(this.WhereToPutTheFilesButton_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.EvaluationFileTemplateName);
            this.groupBox4.Controls.Add(this.FindEvaluationFileTemplateButton);
            this.groupBox4.Location = new System.Drawing.Point(6, 38);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(376, 128);
            this.groupBox4.TabIndex = 15;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Step 1";
            // 
            // EvaluationFileTemplateName
            // 
            this.EvaluationFileTemplateName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.EvaluationFileTemplateName.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.EvaluationFileTemplateName.Location = new System.Drawing.Point(6, 88);
            this.EvaluationFileTemplateName.Name = "EvaluationFileTemplateName";
            this.EvaluationFileTemplateName.Size = new System.Drawing.Size(364, 24);
            this.EvaluationFileTemplateName.TabIndex = 19;
            this.EvaluationFileTemplateName.Text = "Template Name?";
            // 
            // FindEvaluationFileTemplateButton
            // 
            this.FindEvaluationFileTemplateButton.Enabled = false;
            this.FindEvaluationFileTemplateButton.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.FindEvaluationFileTemplateButton.Location = new System.Drawing.Point(6, 38);
            this.FindEvaluationFileTemplateButton.Name = "FindEvaluationFileTemplateButton";
            this.FindEvaluationFileTemplateButton.Size = new System.Drawing.Size(163, 39);
            this.FindEvaluationFileTemplateButton.TabIndex = 18;
            this.FindEvaluationFileTemplateButton.Text = "Find Evaluation File Template";
            this.FindEvaluationFileTemplateButton.UseVisualStyleBackColor = true;
            this.FindEvaluationFileTemplateButton.Click += new System.EventHandler(this.FindEvaluationFileTemplateButton_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.ErrorMessageLabel);
            this.groupBox3.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBox3.Location = new System.Drawing.Point(7, 624);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(790, 74);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Error Message";
            // 
            // ErrorMessageLabel
            // 
            this.ErrorMessageLabel.AutoSize = true;
            this.ErrorMessageLabel.Location = new System.Drawing.Point(12, 24);
            this.ErrorMessageLabel.Name = "ErrorMessageLabel";
            this.ErrorMessageLabel.Size = new System.Drawing.Size(127, 17);
            this.ErrorMessageLabel.TabIndex = 0;
            this.ErrorMessageLabel.Text = "Error Message Here";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(904, 710);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.GetExelFile);
            this.Controls.Add(this.comboBox1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private RichTextBox richTextBox1;
        private Button NextStudentButton;
        private ComboBox comboBox1;
        private Button PreviousStudentButton;
        private Button CreateFilesForEvaluationButton;
        private Button GetExelFile;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label OutHowManyStudents;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private Button WhereToPutTheFilesButton;
        private GroupBox groupBox3;
        private Label ErrorMessageLabel;
        private Button FindEvaluationFileTemplateButton;
        private Label EvaluationFileTemplateName;
        private GroupBox groupBox4;
        private GroupBox groupBox5;
        private Label FolderLocationName;
        private GroupBox groupBox6;
    }
}