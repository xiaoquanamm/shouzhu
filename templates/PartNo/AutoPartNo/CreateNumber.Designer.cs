namespace Interop.Office.Core
{
    partial class CreateNumber
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CreateNumber));
            this.btnOK = new System.Windows.Forms.Button();
            this.cmbStartIndex = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnReSort = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoTemp4 = new System.Windows.Forms.RadioButton();
            this.cmbPPtName = new System.Windows.Forms.ComboBox();
            this.rdoTemp3 = new System.Windows.Forms.RadioButton();
            this.rdoTemp1 = new System.Windows.Forms.RadioButton();
            this.rdoTemp2 = new System.Windows.Forms.RadioButton();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(123, 116);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(87, 23);
            this.btnOK.TabIndex = 19;
            this.btnOK.Text = "生成零件号";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // cmbStartIndex
            // 
            this.cmbStartIndex.DropDownWidth = 200;
            this.cmbStartIndex.FormattingEnabled = true;
            this.cmbStartIndex.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "20",
            "30",
            "40",
            "50",
            "100",
            "101",
            "$项目数×$图号",
            "$属性名（$项目数）",
            "$物料编码 [$项目数]",
            "\"$代号\"\"$名称\"",
            "$图号-$页码-$页数"});
            this.cmbStartIndex.Location = new System.Drawing.Point(13, 42);
            this.cmbStartIndex.Name = "cmbStartIndex";
            this.cmbStartIndex.Size = new System.Drawing.Size(104, 20);
            this.cmbStartIndex.TabIndex = 21;
            this.cmbStartIndex.Text = "1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 22;
            this.label1.Text = "指定起始号：";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnReSort);
            this.groupBox2.Controls.Add(this.cmbStartIndex);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(11, 159);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(216, 76);
            this.groupBox2.TabIndex = 24;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "重排序零件号";
            // 
            // btnReSort
            // 
            this.btnReSort.Location = new System.Drawing.Point(123, 41);
            this.btnReSort.Name = "btnReSort";
            this.btnReSort.Size = new System.Drawing.Size(87, 23);
            this.btnReSort.TabIndex = 25;
            this.btnReSort.Text = "开始重排序";
            this.btnReSort.UseVisualStyleBackColor = true;
            this.btnReSort.Click += new System.EventHandler(this.btnReSort_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Controls.Add(this.rdoTemp4);
            this.groupBox1.Controls.Add(this.cmbPPtName);
            this.groupBox1.Controls.Add(this.rdoTemp3);
            this.groupBox1.Controls.Add(this.rdoTemp1);
            this.groupBox1.Controls.Add(this.rdoTemp2);
            this.groupBox1.Location = new System.Drawing.Point(11, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(216, 145);
            this.groupBox1.TabIndex = 25;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "指定序号";
            // 
            // rdoTemp4
            // 
            this.rdoTemp4.AutoSize = true;
            this.rdoTemp4.Location = new System.Drawing.Point(13, 96);
            this.rdoTemp4.Name = "rdoTemp4";
            this.rdoTemp4.Size = new System.Drawing.Size(107, 16);
            this.rdoTemp4.TabIndex = 6;
            this.rdoTemp4.TabStop = true;
            this.rdoTemp4.Text = "序号为[项目数]";
            this.rdoTemp4.UseVisualStyleBackColor = true;
            // 
            // cmbPPtName
            // 
            this.cmbPPtName.FormattingEnabled = true;
            this.cmbPPtName.Items.AddRange(new object[] {
            "代号",
            "名称",
            "规格",
            "图号"});
            this.cmbPPtName.Location = new System.Drawing.Point(123, 46);
            this.cmbPPtName.Name = "cmbPPtName";
            this.cmbPPtName.Size = new System.Drawing.Size(87, 20);
            this.cmbPPtName.TabIndex = 5;
            this.cmbPPtName.Text = "序号";
            // 
            // rdoTemp3
            // 
            this.rdoTemp3.AutoSize = true;
            this.rdoTemp3.Location = new System.Drawing.Point(13, 72);
            this.rdoTemp3.Name = "rdoTemp3";
            this.rdoTemp3.Size = new System.Drawing.Size(83, 16);
            this.rdoTemp3.TabIndex = 4;
            this.rdoTemp3.TabStop = true;
            this.rdoTemp3.Text = "序号为文字";
            this.rdoTemp3.UseVisualStyleBackColor = true;
            // 
            // rdoTemp1
            // 
            this.rdoTemp1.AutoSize = true;
            this.rdoTemp1.Location = new System.Drawing.Point(13, 24);
            this.rdoTemp1.Name = "rdoTemp1";
            this.rdoTemp1.Size = new System.Drawing.Size(107, 16);
            this.rdoTemp1.TabIndex = 3;
            this.rdoTemp1.TabStop = true;
            this.rdoTemp1.Text = "序号为[项目号]";
            this.rdoTemp1.UseVisualStyleBackColor = true;
            // 
            // rdoTemp2
            // 
            this.rdoTemp2.AutoSize = true;
            this.rdoTemp2.Location = new System.Drawing.Point(13, 48);
            this.rdoTemp2.Name = "rdoTemp2";
            this.rdoTemp2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.rdoTemp2.Size = new System.Drawing.Size(119, 16);
            this.rdoTemp2.TabIndex = 2;
            this.rdoTemp2.TabStop = true;
            this.rdoTemp2.Text = "序号为零件属性：";
            this.rdoTemp2.UseVisualStyleBackColor = true;
            // 
            // CreateNumber
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(239, 242);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CreateNumber";
            this.ShowIcon = true;
            this.ShowInTaskbar = true;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Tag = "1";
            this.Text = "生成零件号";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CreateNumber_FormClosing);
            this.Load += new System.EventHandler(this.CreateNumber_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.ComboBox cmbStartIndex;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnReSort;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoTemp4;
        private System.Windows.Forms.ComboBox cmbPPtName;
        private System.Windows.Forms.RadioButton rdoTemp3;
        private System.Windows.Forms.RadioButton rdoTemp1;
        private System.Windows.Forms.RadioButton rdoTemp2;
    }
}