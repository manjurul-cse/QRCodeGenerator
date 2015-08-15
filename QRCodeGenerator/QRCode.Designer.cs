namespace QRCodeGenerator
{
    partial class QRCode
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(QRCode));
            this.excelFileGetButton = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // excelFileGetButton
            // 
            this.excelFileGetButton.Location = new System.Drawing.Point(12, 12);
            this.excelFileGetButton.Name = "excelFileGetButton";
            this.excelFileGetButton.Size = new System.Drawing.Size(98, 38);
            this.excelFileGetButton.TabIndex = 0;
            this.excelFileGetButton.Text = "Browse";
            this.excelFileGetButton.UseVisualStyleBackColor = true;
            this.excelFileGetButton.Click += new System.EventHandler(this.excelFileGetButton_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 56);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(740, 455);
            this.dataGridView1.TabIndex = 1;
            // 
            // QRCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(764, 523);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.excelFileGetButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "QRCode";
            this.Text = "QRCode";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button excelFileGetButton;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

