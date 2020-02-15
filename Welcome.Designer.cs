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
            this.namNghiepVu = new System.Windows.Forms.DateTimePicker();
            this.start = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // namNghiepVu
            // 
            this.namNghiepVu.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.namNghiepVu.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.namNghiepVu.CustomFormat = "yyyy";
            this.namNghiepVu.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.namNghiepVu.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.namNghiepVu.Location = new System.Drawing.Point(47, 36);
            this.namNghiepVu.Name = "namNghiepVu";
            this.namNghiepVu.Size = new System.Drawing.Size(200, 26);
            this.namNghiepVu.TabIndex = 0;
            this.namNghiepVu.Value = new System.DateTime(2020, 1, 1, 0, 0, 0, 0);
            // 
            // start
            // 
            this.start.Location = new System.Drawing.Point(104, 70);
            this.start.Name = "start";
            this.start.Size = new System.Drawing.Size(79, 37);
            this.start.TabIndex = 1;
            this.start.Text = "Bắt đầu";
            this.start.UseVisualStyleBackColor = true;
            this.start.Click += new System.EventHandler(this.Start_Click);
            // 
            // Welcome
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(279, 119);
            this.Controls.Add(this.start);
            this.Controls.Add(this.namNghiepVu);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Welcome";
            this.Text = "Đối chiếu công nợ";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DateTimePicker namNghiepVu;
        private System.Windows.Forms.Button start;
    }
}