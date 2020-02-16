namespace CongNo
{
    partial class Welcome
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Welcome));
            this.start = new System.Windows.Forms.Button();
            this.numericNam = new System.Windows.Forms.NumericUpDown();
            this.comboNam = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.numericNam)).BeginInit();
            this.SuspendLayout();
            // 
            // start
            // 
            this.start.Location = new System.Drawing.Point(46, 70);
            this.start.Name = "start";
            this.start.Size = new System.Drawing.Size(81, 37);
            this.start.TabIndex = 1;
            this.start.Text = "Bắt đầu";
            this.start.UseVisualStyleBackColor = true;
            this.start.Click += new System.EventHandler(this.Start_Click);
            // 
            // numericNam
            // 
            this.numericNam.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numericNam.Location = new System.Drawing.Point(46, 30);
            this.numericNam.Maximum = new decimal(new int[] {
            2050,
            0,
            0,
            0});
            this.numericNam.Minimum = new decimal(new int[] {
            2019,
            0,
            0,
            0});
            this.numericNam.Name = "numericNam";
            this.numericNam.Size = new System.Drawing.Size(81, 22);
            this.numericNam.TabIndex = 2;
            this.numericNam.TabStop = false;
            this.numericNam.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.numericNam.Value = new decimal(new int[] {
            2019,
            0,
            0,
            0});
            // 
            // comboNam
            // 
            this.comboNam.FormattingEnabled = true;
            this.comboNam.Location = new System.Drawing.Point(46, 30);
            this.comboNam.Name = "comboNam";
            this.comboNam.Size = new System.Drawing.Size(81, 21);
            this.comboNam.TabIndex = 3;
            // 
            // Welcome
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(172, 119);
            this.Controls.Add(this.comboNam);
            this.Controls.Add(this.numericNam);
            this.Controls.Add(this.start);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Welcome";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Đối chiếu công nợ";
            this.Load += new System.EventHandler(this.Welcome_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numericNam)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button start;
        private System.Windows.Forms.NumericUpDown numericNam;
        private System.Windows.Forms.ComboBox comboNam;
    }
}