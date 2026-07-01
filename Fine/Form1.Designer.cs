namespace Fine
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
            components = new System.ComponentModel.Container();
            System.Windows.Forms.Timer timer1;
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            txtPlate = new TextBox();
            txtTime = new TextBox();
            label1 = new Label();
            label2 = new Label();
            groupBox1 = new GroupBox();
            label5 = new Label();
            label3 = new Label();
            label4 = new Label();
            btnGenerate = new Button();
            txtSeq = new TextBox();
            label6 = new Label();
            picPreview1 = new PictureBox();
            picPreview2 = new PictureBox();
            lblPreview1 = new Label();
            lblPreview2 = new Label();
            btnShortcut = new Button();
            button5 = new Button();
            button4 = new Button();
            btnMatchAdmin = new Button();
            btnOpenExcel = new Button();
            btnMatchExternal = new Button();
            btnHelp = new Button();
            btnSingleMatch = new Button();
            btnBatchOcrMatch = new Button();
            lblViolationType = new Label();
            cboViolationType = new ComboBox();
            lblImageQueueStatus = new Label();
            btnClearQueue = new Button();
            timer1 = new System.Windows.Forms.Timer(components);
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)picPreview1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)picPreview2).BeginInit();
            SuspendLayout();
            // 
            // timer1
            // 
            timer1.Enabled = true;
            timer1.Interval = 500;
            timer1.Tick += timer1_Tick;
            // 
            // button1
            // 
            button1.Location = new Point(738, 236);
            button1.Name = "button1";
            button1.Size = new Size(156, 34);
            button1.TabIndex = 0;
            button1.Text = "读取剪贴板文字";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(915, 236);
            button2.Name = "button2";
            button2.Size = new Size(153, 34);
            button2.TabIndex = 1;
            button2.Text = "导入模板";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Location = new Point(915, 297);
            button3.Name = "button3";
            button3.Size = new Size(172, 34);
            button3.TabIndex = 2;
            button3.Text = "WORD转PDF";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // txtPlate
            // 
            txtPlate.Location = new Point(169, 55);
            txtPlate.Name = "txtPlate";
            txtPlate.Size = new Size(243, 30);
            txtPlate.TabIndex = 3;
            // 
            // txtTime
            // 
            txtTime.Location = new Point(167, 91);
            txtTime.Name = "txtTime";
            txtTime.Size = new Size(245, 30);
            txtTime.TabIndex = 4;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(39, 61);
            label1.Name = "label1";
            label1.Size = new Size(82, 24);
            label1.TabIndex = 5;
            label1.Text = "车牌号：";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(39, 99);
            label2.Name = "label2";
            label2.Size = new Size(100, 24);
            label2.TabIndex = 6;
            label2.Text = "违规时间：";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(label5);
            groupBox1.Controls.Add(label3);
            groupBox1.Controls.Add(label4);
            groupBox1.Location = new Point(714, 2);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(423, 218);
            groupBox1.TabIndex = 8;
            groupBox1.TabStop = false;
            groupBox1.Text = "操作指引 (工作流)";
            // 
            // label5
            // 
            label5.Location = new Point(8, 147);
            label5.Name = "label5";
            label5.Size = new Size(431, 36);
            label5.TabIndex = 2;
            label5.Text = "生成：点击[确认生成] 或 按 Shift + Enter。";
            // 
            // label3
            // 
            label3.Location = new Point(8, 94);
            label3.Name = "label3";
            label3.Size = new Size(409, 53);
            label3.TabIndex = 1;
            label3.Text = "抓图片：再次截图违规轨迹 -> 点击微信截图的[√] (复制到剪贴板)";
            // 
            // label4
            // 
            label4.Location = new Point(11, 34);
            label4.Name = "label4";
            label4.Size = new Size(395, 49);
            label4.TabIndex = 0;
            label4.Text = "抓文字：微信截图(Alt+A) -> 提取文字 -> 复制(Ctrl+C) -> 点击[读取文字]";
            // 
            // btnGenerate
            // 
            btnGenerate.Location = new Point(738, 297);
            btnGenerate.Name = "btnGenerate";
            btnGenerate.Size = new Size(152, 34);
            btnGenerate.TabIndex = 9;
            btnGenerate.Text = "确认生成处罚单";
            btnGenerate.UseVisualStyleBackColor = true;
            btnGenerate.Click += btnGenerate_Click;
            // 
            // txtSeq
            // 
            txtSeq.Location = new Point(171, 12);
            txtSeq.Name = "txtSeq";
            txtSeq.Size = new Size(150, 30);
            txtSeq.TabIndex = 10;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(39, 19);
            label6.Name = "label6";
            label6.Size = new Size(82, 24);
            label6.TabIndex = 11;
            label6.Text = "单号结尾";
            // 
            // lblPreview1
            // 
            lblPreview1.AutoSize = true;
            lblPreview1.Location = new Point(10, 174);
            lblPreview1.Name = "lblPreview1";
            lblPreview1.Size = new Size(80, 24);
            lblPreview1.TabIndex = 26;
            lblPreview1.Text = "证据图 1";
            // 
            // lblPreview2
            // 
            lblPreview2.AutoSize = true;
            lblPreview2.Location = new Point(400, 174);
            lblPreview2.Name = "lblPreview2";
            lblPreview2.Size = new Size(80, 24);
            lblPreview2.TabIndex = 27;
            lblPreview2.Text = "证据图 2";
            // 
            // picPreview1
            // 
            picPreview1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            picPreview1.BorderStyle = BorderStyle.FixedSingle;
            picPreview1.Location = new Point(-3, 200);
            picPreview1.Name = "picPreview1";
            picPreview1.Size = new Size(390, 610);
            picPreview1.SizeMode = PictureBoxSizeMode.Zoom;
            picPreview1.TabIndex = 12;
            picPreview1.TabStop = false;
            picPreview1.Click += pictureBox1_Click;
            // 
            // picPreview2
            // 
            picPreview2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            picPreview2.BorderStyle = BorderStyle.FixedSingle;
            picPreview2.Location = new Point(395, 200);
            picPreview2.Name = "picPreview2";
            picPreview2.Size = new Size(298, 610);
            picPreview2.SizeMode = PictureBoxSizeMode.Zoom;
            picPreview2.TabIndex = 28;
            picPreview2.TabStop = false;
            // 
            // lblImageQueueStatus
            // 
            lblImageQueueStatus.AutoSize = true;
            lblImageQueueStatus.ForeColor = System.Drawing.Color.DarkBlue;
            lblImageQueueStatus.Location = new Point(3, 818);
            lblImageQueueStatus.Name = "lblImageQueueStatus";
            lblImageQueueStatus.Size = new Size(150, 24);
            lblImageQueueStatus.TabIndex = 24;
            lblImageQueueStatus.Text = "已选 0/2 张证据图";
            // 
            // btnClearQueue
            // 
            btnClearQueue.Location = new Point(200, 814);
            btnClearQueue.Name = "btnClearQueue";
            btnClearQueue.Size = new Size(100, 34);
            btnClearQueue.TabIndex = 25;
            btnClearQueue.Text = "清空图片";
            btnClearQueue.UseVisualStyleBackColor = true;
            btnClearQueue.Click += btnClearQueue_Click;
            // 
            // btnShortcut
            // 
            btnShortcut.Location = new Point(738, 356);
            btnShortcut.Name = "btnShortcut";
            btnShortcut.Size = new Size(148, 34);
            btnShortcut.TabIndex = 13;
            btnShortcut.Text = "添加快捷方式";
            btnShortcut.UseVisualStyleBackColor = true;
            btnShortcut.Click += btnShortcut_Click;
            // 
            // button5
            // 
            button5.Location = new Point(738, 406);
            button5.Name = "button5";
            button5.Size = new Size(148, 34);
            button5.TabIndex = 14;
            button5.Text = "生成群通报";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // button4
            // 
            button4.Location = new Point(743, 459);
            button4.Name = "button4";
            button4.Size = new Size(143, 34);
            button4.TabIndex = 15;
            button4.Text = "生成压缩包";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // btnMatchAdmin
            // 
            btnMatchAdmin.Location = new Point(912, 356);
            btnMatchAdmin.Name = "btnMatchAdmin";
            btnMatchAdmin.Size = new Size(175, 34);
            btnMatchAdmin.TabIndex = 16;
            btnMatchAdmin.Text = "内部详细匹配";
            btnMatchAdmin.UseVisualStyleBackColor = true;
            btnMatchAdmin.Click += btnMatchAdmin_Click;
            // 
            // btnOpenExcel
            // 
            btnOpenExcel.Location = new Point(915, 406);
            btnOpenExcel.Name = "btnOpenExcel";
            btnOpenExcel.Size = new Size(153, 34);
            btnOpenExcel.TabIndex = 17;
            btnOpenExcel.Text = "打开本地台账";
            btnOpenExcel.UseVisualStyleBackColor = true;
            btnOpenExcel.Click += btnOpenExcel_Click;
            // 
            // btnMatchExternal
            // 
            btnMatchExternal.Location = new Point(915, 459);
            btnMatchExternal.Name = "btnMatchExternal";
            btnMatchExternal.Size = new Size(189, 34);
            btnMatchExternal.TabIndex = 18;
            btnMatchExternal.Text = "生成外部通报清单";
            btnMatchExternal.UseVisualStyleBackColor = true;
            btnMatchExternal.Click += btnMatchExternal_Click;
            // 
            // btnHelp
            // 
            btnHelp.Location = new Point(545, 12);
            btnHelp.Name = "btnHelp";
            btnHelp.Size = new Size(112, 34);
            btnHelp.TabIndex = 19;
            btnHelp.Text = "使用帮助";
            btnHelp.UseVisualStyleBackColor = true;
            btnHelp.Click += btnHelp_Click;
            // 
            // btnSingleMatch
            // 
            btnSingleMatch.Location = new Point(744, 520);
            btnSingleMatch.Name = "btnSingleMatch";
            btnSingleMatch.Size = new Size(112, 34);
            btnSingleMatch.TabIndex = 20;
            btnSingleMatch.Text = "单车匹配";
            btnSingleMatch.UseVisualStyleBackColor = true;
            btnSingleMatch.Click += btnSingleMatch_Click;
            // 
            // btnBatchOcrMatch
            // 
            btnBatchOcrMatch.Location = new Point(915, 520);
            btnBatchOcrMatch.Name = "btnBatchOcrMatch";
            btnBatchOcrMatch.Size = new Size(180, 34);
            btnBatchOcrMatch.TabIndex = 21;
            btnBatchOcrMatch.Text = "批量截图识别匹配";
            btnBatchOcrMatch.UseVisualStyleBackColor = true;
            btnBatchOcrMatch.Click += btnBatchOcrMatch_Click;
            // 
            // lblViolationType
            // 
            lblViolationType.AutoSize = true;
            lblViolationType.Location = new Point(39, 135);
            lblViolationType.Name = "lblViolationType";
            lblViolationType.Size = new Size(100, 24);
            lblViolationType.TabIndex = 22;
            lblViolationType.Text = "违规类型：";
            // 
            // cboViolationType
            // 
            cboViolationType.DropDownStyle = ComboBoxStyle.DropDownList;
            cboViolationType.Location = new Point(169, 131);
            cboViolationType.Name = "cboViolationType";
            cboViolationType.Size = new Size(243, 32);
            cboViolationType.TabIndex = 23;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(11F, 24F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            ClientSize = new Size(1149, 860);
            Controls.Add(btnBatchOcrMatch);
            Controls.Add(cboViolationType);
            Controls.Add(lblViolationType);
            Controls.Add(lblImageQueueStatus);
            Controls.Add(btnClearQueue);
            Controls.Add(btnSingleMatch);
            Controls.Add(btnHelp);
            Controls.Add(btnMatchExternal);
            Controls.Add(btnOpenExcel);
            Controls.Add(btnMatchAdmin);
            Controls.Add(button4);
            Controls.Add(button5);
            Controls.Add(btnShortcut);
            Controls.Add(picPreview2);
            Controls.Add(picPreview1);
            Controls.Add(lblPreview2);
            Controls.Add(lblPreview1);
            Controls.Add(label6);
            Controls.Add(txtSeq);
            Controls.Add(btnGenerate);
            Controls.Add(groupBox1);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(txtTime);
            Controls.Add(txtPlate);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)picPreview1).EndInit();
            ((System.ComponentModel.ISupportInitialize)picPreview2).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private Button button3;
        private TextBox txtPlate;
        private TextBox txtTime;
        private Label label1;
        private Label label2;
        private GroupBox groupBox1;
        private Label label5;
        private Label label3;
        private Label label4;
        private Button btnGenerate;
        private TextBox txtSeq;
        private Label label6;
        private PictureBox picPreview1;
        private PictureBox picPreview2;
        private Label lblPreview1;
        private Label lblPreview2;
        private System.Windows.Forms.Timer timer1;
        private Button btnShortcut;
        private Button button5;
        private Button button4;
        private Button btnMatchAdmin;
        private Button btnOpenExcel;
        private Button btnMatchExternal;
        private Button btnHelp;
        private Button btnSingleMatch;
        private Button btnBatchOcrMatch;
        private Label lblViolationType;
        private ComboBox cboViolationType;
        private Label lblImageQueueStatus;
        private Button btnClearQueue;
    }
}
