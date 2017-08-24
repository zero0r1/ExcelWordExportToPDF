namespace WindowsFormsTestApplication
{
    partial class TestForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TestForm));
            this.label1 = new System.Windows.Forms.Label();
            this.labSourceExcel = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.labSourceWord = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnWord = new System.Windows.Forms.Button();
            this.btnConvertPDF = new System.Windows.Forms.Button();
            this.btn = new System.Windows.Forms.Button();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "File Address:";
            // 
            // labSourceExcel
            // 
            this.labSourceExcel.AutoSize = true;
            this.labSourceExcel.Location = new System.Drawing.Point(78, 13);
            this.labSourceExcel.Name = "labSourceExcel";
            this.labSourceExcel.Size = new System.Drawing.Size(0, 12);
            this.labSourceExcel.TabIndex = 1;
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(555, 8);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(98, 23);
            this.btnExcel.TabIndex = 2;
            this.btnExcel.Text = "Choose Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // labSourceWord
            // 
            this.labSourceWord.AutoSize = true;
            this.labSourceWord.Location = new System.Drawing.Point(78, 39);
            this.labSourceWord.Name = "labSourceWord";
            this.labSourceWord.Size = new System.Drawing.Size(0, 12);
            this.labSourceWord.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 12);
            this.label3.TabIndex = 3;
            this.label3.Text = "File Address:";
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(455, 10);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(98, 23);
            this.btnWord.TabIndex = 5;
            this.btnWord.Text = "Choose Word";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // btnConvertPDF
            // 
            this.btnConvertPDF.Location = new System.Drawing.Point(555, 39);
            this.btnConvertPDF.Name = "btnConvertPDF";
            this.btnConvertPDF.Size = new System.Drawing.Size(98, 23);
            this.btnConvertPDF.TabIndex = 6;
            this.btnConvertPDF.Text = "Excel To PDF";
            this.btnConvertPDF.UseVisualStyleBackColor = true;
            this.btnConvertPDF.Click += new System.EventHandler(this.btnConvertPDF_Click);
            // 
            // btn
            // 
            this.btn.Location = new System.Drawing.Point(455, 39);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(98, 23);
            this.btn.TabIndex = 7;
            this.btn.Text = "Word To PDF";
            this.btn.UseVisualStyleBackColor = true;
            this.btn.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(15, 79);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(432, 180);
            this.checkedListBox1.TabIndex = 8;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(555, 68);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(98, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "PDF To Each";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(555, 97);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(98, 23);
            this.button2.TabIndex = 10;
            this.button2.Text = "PDF To Marge";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // TestForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(661, 268);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.btn);
            this.Controls.Add(this.btnConvertPDF);
            this.Controls.Add(this.btnWord);
            this.Controls.Add(this.labSourceWord);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.labSourceExcel);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TestForm";
            this.Text = "TestForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labSourceExcel;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label labSourceWord;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Button btnConvertPDF;
        private System.Windows.Forms.Button btn;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

