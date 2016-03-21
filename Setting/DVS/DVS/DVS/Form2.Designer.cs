namespace DVS
{
    partial class Form2
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
            this.components = new System.ComponentModel.Container();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.autoSetPannel = new System.Windows.Forms.Panel();
            this.autoCancerBt = new System.Windows.Forms.Button();
            this.autoSaveBt = new System.Windows.Forms.Button();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.autoSetLabel02 = new System.Windows.Forms.Label();
            this.autoSetLabel01 = new System.Windows.Forms.Label();
            this.deskPannel = new System.Windows.Forms.Panel();
            this.manualSetPannel = new System.Windows.Forms.Panel();
            this.manualRun04 = new System.Windows.Forms.Button();
            this.manualRun03 = new System.Windows.Forms.Button();
            this.manualRun02 = new System.Windows.Forms.Button();
            this.manualRun01 = new System.Windows.Forms.Button();
            this.manualSetLabelStr04 = new System.Windows.Forms.Label();
            this.manualSetLabelStr02 = new System.Windows.Forms.Label();
            this.manualSetLabelStr03 = new System.Windows.Forms.Label();
            this.manualSetLabelStr01 = new System.Windows.Forms.Label();
            this.manualClosoe = new System.Windows.Forms.Button();
            this.manualSetLabel01 = new System.Windows.Forms.Label();
            this.copyDataPanel = new System.Windows.Forms.Panel();
            this.copyDataLabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.autoSetPannel.SuspendLayout();
            this.manualSetPannel.SuspendLayout();
            this.copyDataPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(472, 20);
            this.panel1.TabIndex = 2;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(419, 4);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 12);
            this.label4.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(335, 4);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "自動匯入時間";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(68, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 12);
            this.label2.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "伺服器時間";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.button2);
            this.panel3.Controls.Add(this.button1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 265);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(472, 28);
            this.panel3.TabIndex = 4;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(376, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(93, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "自動匯入設定";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(295, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 11;
            this.button1.Text = "手動匯入";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // autoSetPannel
            // 
            this.autoSetPannel.BackColor = System.Drawing.Color.LightYellow;
            this.autoSetPannel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.autoSetPannel.Controls.Add(this.autoCancerBt);
            this.autoSetPannel.Controls.Add(this.autoSaveBt);
            this.autoSetPannel.Controls.Add(this.comboBox3);
            this.autoSetPannel.Controls.Add(this.label8);
            this.autoSetPannel.Controls.Add(this.comboBox2);
            this.autoSetPannel.Controls.Add(this.label7);
            this.autoSetPannel.Controls.Add(this.comboBox1);
            this.autoSetPannel.Controls.Add(this.autoSetLabel02);
            this.autoSetPannel.Controls.Add(this.autoSetLabel01);
            this.autoSetPannel.Location = new System.Drawing.Point(136, 66);
            this.autoSetPannel.Name = "autoSetPannel";
            this.autoSetPannel.Size = new System.Drawing.Size(200, 93);
            this.autoSetPannel.TabIndex = 5;
            this.autoSetPannel.Visible = false;
            // 
            // autoCancerBt
            // 
            this.autoCancerBt.BackColor = System.Drawing.SystemColors.Control;
            this.autoCancerBt.ForeColor = System.Drawing.Color.Red;
            this.autoCancerBt.Location = new System.Drawing.Point(21, 54);
            this.autoCancerBt.Name = "autoCancerBt";
            this.autoCancerBt.Size = new System.Drawing.Size(75, 23);
            this.autoCancerBt.TabIndex = 8;
            this.autoCancerBt.Text = "關閉設定";
            this.autoCancerBt.UseVisualStyleBackColor = false;
            this.autoCancerBt.Click += new System.EventHandler(this.autoBt_Click);
            // 
            // autoSaveBt
            // 
            this.autoSaveBt.BackColor = System.Drawing.SystemColors.Control;
            this.autoSaveBt.ForeColor = System.Drawing.Color.Blue;
            this.autoSaveBt.Location = new System.Drawing.Point(102, 54);
            this.autoSaveBt.Name = "autoSaveBt";
            this.autoSaveBt.Size = new System.Drawing.Size(75, 23);
            this.autoSaveBt.TabIndex = 7;
            this.autoSaveBt.Text = "儲存設定";
            this.autoSaveBt.UseVisualStyleBackColor = false;
            this.autoSaveBt.Click += new System.EventHandler(this.autoBt_Click);
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Location = new System.Drawing.Point(148, 28);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(40, 20);
            this.comboBox3.TabIndex = 6;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label8.Location = new System.Drawing.Point(134, 32);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(8, 12);
            this.label8.TabIndex = 5;
            this.label8.Text = ":";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(93, 28);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(40, 20);
            this.comboBox2.TabIndex = 4;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label7.Location = new System.Drawing.Point(79, 31);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(8, 12);
            this.label7.TabIndex = 3;
            this.label7.Text = ":";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(38, 28);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(40, 20);
            this.comboBox1.TabIndex = 2;
            // 
            // autoSetLabel02
            // 
            this.autoSetLabel02.AutoSize = true;
            this.autoSetLabel02.Location = new System.Drawing.Point(3, 32);
            this.autoSetLabel02.Name = "autoSetLabel02";
            this.autoSetLabel02.Size = new System.Drawing.Size(29, 12);
            this.autoSetLabel02.TabIndex = 1;
            this.autoSetLabel02.Text = "時間";
            // 
            // autoSetLabel01
            // 
            this.autoSetLabel01.Dock = System.Windows.Forms.DockStyle.Top;
            this.autoSetLabel01.Location = new System.Drawing.Point(0, 0);
            this.autoSetLabel01.Name = "autoSetLabel01";
            this.autoSetLabel01.Size = new System.Drawing.Size(198, 20);
            this.autoSetLabel01.TabIndex = 0;
            this.autoSetLabel01.Text = "自動匯入設定";
            this.autoSetLabel01.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // deskPannel
            // 
            this.deskPannel.BackColor = System.Drawing.Color.CornflowerBlue;
            this.deskPannel.Location = new System.Drawing.Point(0, 22);
            this.deskPannel.Name = "deskPannel";
            this.deskPannel.Size = new System.Drawing.Size(24, 20);
            this.deskPannel.TabIndex = 7;
            // 
            // manualSetPannel
            // 
            this.manualSetPannel.BackColor = System.Drawing.Color.Khaki;
            this.manualSetPannel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.manualSetPannel.Controls.Add(this.manualRun04);
            this.manualSetPannel.Controls.Add(this.manualRun03);
            this.manualSetPannel.Controls.Add(this.manualRun02);
            this.manualSetPannel.Controls.Add(this.manualRun01);
            this.manualSetPannel.Controls.Add(this.manualSetLabelStr04);
            this.manualSetPannel.Controls.Add(this.manualSetLabelStr02);
            this.manualSetPannel.Controls.Add(this.manualSetLabelStr03);
            this.manualSetPannel.Controls.Add(this.manualSetLabelStr01);
            this.manualSetPannel.Controls.Add(this.manualClosoe);
            this.manualSetPannel.Controls.Add(this.manualSetLabel01);
            this.manualSetPannel.Location = new System.Drawing.Point(110, 69);
            this.manualSetPannel.Name = "manualSetPannel";
            this.manualSetPannel.Size = new System.Drawing.Size(260, 154);
            this.manualSetPannel.TabIndex = 9;
            this.manualSetPannel.Visible = false;
            // 
            // manualRun04
            // 
            this.manualRun04.BackColor = System.Drawing.SystemColors.Control;
            this.manualRun04.ForeColor = System.Drawing.Color.Blue;
            this.manualRun04.Location = new System.Drawing.Point(177, 98);
            this.manualRun04.Name = "manualRun04";
            this.manualRun04.Size = new System.Drawing.Size(44, 23);
            this.manualRun04.TabIndex = 24;
            this.manualRun04.Text = "執行";
            this.manualRun04.UseVisualStyleBackColor = false;
            this.manualRun04.Click += new System.EventHandler(this.manualRun_Click);
            // 
            // manualRun03
            // 
            this.manualRun03.BackColor = System.Drawing.SystemColors.Control;
            this.manualRun03.Enabled = false;
            this.manualRun03.ForeColor = System.Drawing.Color.Blue;
            this.manualRun03.Location = new System.Drawing.Point(177, 72);
            this.manualRun03.Name = "manualRun03";
            this.manualRun03.Size = new System.Drawing.Size(44, 23);
            this.manualRun03.TabIndex = 23;
            this.manualRun03.Text = "執行";
            this.manualRun03.UseVisualStyleBackColor = false;
            this.manualRun03.Click += new System.EventHandler(this.manualRun_Click);
            // 
            // manualRun02
            // 
            this.manualRun02.BackColor = System.Drawing.SystemColors.Control;
            this.manualRun02.Enabled = false;
            this.manualRun02.ForeColor = System.Drawing.Color.Blue;
            this.manualRun02.Location = new System.Drawing.Point(177, 46);
            this.manualRun02.Name = "manualRun02";
            this.manualRun02.Size = new System.Drawing.Size(44, 23);
            this.manualRun02.TabIndex = 22;
            this.manualRun02.Text = "執行";
            this.manualRun02.UseVisualStyleBackColor = false;
            this.manualRun02.Click += new System.EventHandler(this.manualRun_Click);
            // 
            // manualRun01
            // 
            this.manualRun01.BackColor = System.Drawing.SystemColors.Control;
            this.manualRun01.Enabled = false;
            this.manualRun01.ForeColor = System.Drawing.Color.Blue;
            this.manualRun01.Location = new System.Drawing.Point(177, 20);
            this.manualRun01.Name = "manualRun01";
            this.manualRun01.Size = new System.Drawing.Size(44, 23);
            this.manualRun01.TabIndex = 21;
            this.manualRun01.Text = "執行";
            this.manualRun01.UseVisualStyleBackColor = false;
            this.manualRun01.Click += new System.EventHandler(this.manualRun_Click);
            // 
            // manualSetLabelStr04
            // 
            this.manualSetLabelStr04.AutoSize = true;
            this.manualSetLabelStr04.Location = new System.Drawing.Point(37, 104);
            this.manualSetLabelStr04.Name = "manualSetLabelStr04";
            this.manualSetLabelStr04.Size = new System.Drawing.Size(137, 12);
            this.manualSetLabelStr04.TabIndex = 15;
            this.manualSetLabelStr04.Text = "重新匯入及刪除補課資訊";
            // 
            // manualSetLabelStr02
            // 
            this.manualSetLabelStr02.AutoSize = true;
            this.manualSetLabelStr02.Location = new System.Drawing.Point(37, 52);
            this.manualSetLabelStr02.Name = "manualSetLabelStr02";
            this.manualSetLabelStr02.Size = new System.Drawing.Size(137, 12);
            this.manualSetLabelStr02.TabIndex = 14;
            this.manualSetLabelStr02.Text = "重新匯入及刪除班級設定";
            // 
            // manualSetLabelStr03
            // 
            this.manualSetLabelStr03.AutoSize = true;
            this.manualSetLabelStr03.Location = new System.Drawing.Point(37, 78);
            this.manualSetLabelStr03.Name = "manualSetLabelStr03";
            this.manualSetLabelStr03.Size = new System.Drawing.Size(137, 12);
            this.manualSetLabelStr03.TabIndex = 13;
            this.manualSetLabelStr03.Text = "重新匯入及刪除教室日誌";
            // 
            // manualSetLabelStr01
            // 
            this.manualSetLabelStr01.AutoSize = true;
            this.manualSetLabelStr01.Location = new System.Drawing.Point(38, 26);
            this.manualSetLabelStr01.Name = "manualSetLabelStr01";
            this.manualSetLabelStr01.Size = new System.Drawing.Size(137, 12);
            this.manualSetLabelStr01.TabIndex = 12;
            this.manualSetLabelStr01.Text = "重新匯入及刪除學生資料";
            // 
            // manualClosoe
            // 
            this.manualClosoe.BackColor = System.Drawing.SystemColors.Control;
            this.manualClosoe.ForeColor = System.Drawing.Color.Red;
            this.manualClosoe.Location = new System.Drawing.Point(92, 123);
            this.manualClosoe.Name = "manualClosoe";
            this.manualClosoe.Size = new System.Drawing.Size(75, 23);
            this.manualClosoe.TabIndex = 10;
            this.manualClosoe.Text = "關閉視窗";
            this.manualClosoe.UseVisualStyleBackColor = false;
            this.manualClosoe.Click += new System.EventHandler(this.manualClosoe_Click);
            // 
            // manualSetLabel01
            // 
            this.manualSetLabel01.Dock = System.Windows.Forms.DockStyle.Top;
            this.manualSetLabel01.Location = new System.Drawing.Point(0, 0);
            this.manualSetLabel01.Name = "manualSetLabel01";
            this.manualSetLabel01.Size = new System.Drawing.Size(258, 20);
            this.manualSetLabel01.TabIndex = 1;
            this.manualSetLabel01.Text = "手動匯入設定";
            this.manualSetLabel01.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // copyDataPanel
            // 
            this.copyDataPanel.BackColor = System.Drawing.Color.Tan;
            this.copyDataPanel.Controls.Add(this.copyDataLabel);
            this.copyDataPanel.Location = new System.Drawing.Point(0, 48);
            this.copyDataPanel.Name = "copyDataPanel";
            this.copyDataPanel.Size = new System.Drawing.Size(100, 82);
            this.copyDataPanel.TabIndex = 11;
            this.copyDataPanel.Visible = false;
            // 
            // copyDataLabel
            // 
            this.copyDataLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.copyDataLabel.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.copyDataLabel.Location = new System.Drawing.Point(0, 0);
            this.copyDataLabel.Name = "copyDataLabel";
            this.copyDataLabel.Size = new System.Drawing.Size(100, 82);
            this.copyDataLabel.TabIndex = 1;
            this.copyDataLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 293);
            this.Controls.Add(this.copyDataPanel);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.manualSetPannel);
            this.Controls.Add(this.autoSetPannel);
            this.Controls.Add(this.deskPannel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form2";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.autoSetPannel.ResumeLayout(false);
            this.autoSetPannel.PerformLayout();
            this.manualSetPannel.ResumeLayout(false);
            this.manualSetPannel.PerformLayout();
            this.copyDataPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel autoSetPannel;
        private System.Windows.Forms.Button autoCancerBt;
        private System.Windows.Forms.Button autoSaveBt;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label autoSetLabel02;
        private System.Windows.Forms.Label autoSetLabel01;
        private System.Windows.Forms.Panel deskPannel;
        private System.Windows.Forms.Panel manualSetPannel;
        private System.Windows.Forms.Label manualSetLabel01;
        private System.Windows.Forms.Button manualClosoe;
        private System.Windows.Forms.Label manualSetLabelStr03;
        private System.Windows.Forms.Label manualSetLabelStr01;
        private System.Windows.Forms.Label manualSetLabelStr02;
        private System.Windows.Forms.Label manualSetLabelStr04;
        private System.Windows.Forms.Button manualRun01;
        private System.Windows.Forms.Button manualRun04;
        private System.Windows.Forms.Button manualRun03;
        private System.Windows.Forms.Button manualRun02;
        private System.Windows.Forms.Panel copyDataPanel;
        private System.Windows.Forms.Label copyDataLabel;
    }
}