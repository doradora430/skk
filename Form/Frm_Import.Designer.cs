namespace SagyouKousuuKanriImport
{
    partial class Frm_Import
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.Txt_FilePath = new System.Windows.Forms.TextBox();
            this.Btn_Refer = new System.Windows.Forms.Button();
            this.Btn_Import = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Txt_FilePath
            // 
            this.Txt_FilePath.Enabled = false;
            this.Txt_FilePath.Location = new System.Drawing.Point(28, 51);
            this.Txt_FilePath.Name = "Txt_FilePath";
            this.Txt_FilePath.Size = new System.Drawing.Size(679, 19);
            this.Txt_FilePath.TabIndex = 0;
            // 
            // Btn_Refer
            // 
            this.Btn_Refer.Location = new System.Drawing.Point(713, 51);
            this.Btn_Refer.Name = "Btn_Refer";
            this.Btn_Refer.Size = new System.Drawing.Size(75, 23);
            this.Btn_Refer.TabIndex = 1;
            this.Btn_Refer.Text = "参照";
            this.Btn_Refer.UseVisualStyleBackColor = true;
            this.Btn_Refer.Click += new System.EventHandler(this.Btn_Refer_Click);
            // 
            // Btn_Import
            // 
            this.Btn_Import.Location = new System.Drawing.Point(713, 107);
            this.Btn_Import.Name = "Btn_Import";
            this.Btn_Import.Size = new System.Drawing.Size(75, 23);
            this.Btn_Import.TabIndex = 2;
            this.Btn_Import.Text = "取込";
            this.Btn_Import.UseVisualStyleBackColor = true;
            this.Btn_Import.Click += new System.EventHandler(this.Btn_Import_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 157);
            this.Controls.Add(this.Btn_Import);
            this.Controls.Add(this.Btn_Refer);
            this.Controls.Add(this.Txt_FilePath);
            this.Name = "Form1";
            this.Text = "作業工数管理インポート";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Txt_FilePath;
        private System.Windows.Forms.Button Btn_Refer;
        private System.Windows.Forms.Button Btn_Import;
    }
}

