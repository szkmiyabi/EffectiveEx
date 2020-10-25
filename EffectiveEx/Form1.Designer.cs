namespace EffectiveEx
{
    partial class Form1
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.openButton = new System.Windows.Forms.Button();
            this.deleteRowButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.sheetNameCombo = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.columnValues = new System.Windows.Forms.TextBox();
            this.reportText = new System.Windows.Forms.TextBox();
            this.flowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.searchValues = new System.Windows.Forms.TextBox();
            this.statusBarControl = new System.Windows.Forms.StatusStrip();
            this.statusText = new System.Windows.Forms.ToolStripStatusLabel();
            this.flowLayoutPanel4 = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.outputDirectory = new System.Windows.Forms.TextBox();
            this.browseOutputDirectoryButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            this.flowLayoutPanel3.SuspendLayout();
            this.statusBarControl.SuspendLayout();
            this.flowLayoutPanel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.reportText, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel3, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel4, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.25373F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.74627F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 116F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 113F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(494, 368);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.openButton);
            this.flowLayoutPanel1.Controls.Add(this.deleteRowButton);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(480, 42);
            this.flowLayoutPanel1.TabIndex = 0;
            // 
            // openButton
            // 
            this.openButton.Location = new System.Drawing.Point(3, 3);
            this.openButton.Name = "openButton";
            this.openButton.Size = new System.Drawing.Size(38, 23);
            this.openButton.TabIndex = 1;
            this.openButton.Text = "開く";
            this.openButton.UseVisualStyleBackColor = true;
            // 
            // deleteRowButton
            // 
            this.deleteRowButton.Location = new System.Drawing.Point(47, 3);
            this.deleteRowButton.Name = "deleteRowButton";
            this.deleteRowButton.Size = new System.Drawing.Size(60, 23);
            this.deleteRowButton.TabIndex = 2;
            this.deleteRowButton.Text = "行削除";
            this.deleteRowButton.UseVisualStyleBackColor = true;
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Controls.Add(this.label2);
            this.flowLayoutPanel2.Controls.Add(this.sheetNameCombo);
            this.flowLayoutPanel2.Controls.Add(this.label4);
            this.flowLayoutPanel2.Controls.Add(this.columnValues);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 90);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(488, 45);
            this.flowLayoutPanel2.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "対象シート";
            // 
            // sheetNameCombo
            // 
            this.sheetNameCombo.FormattingEnabled = true;
            this.sheetNameCombo.Location = new System.Drawing.Point(73, 3);
            this.sheetNameCombo.Name = "sheetNameCombo";
            this.sheetNameCombo.Size = new System.Drawing.Size(121, 21);
            this.sheetNameCombo.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(200, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 13);
            this.label4.TabIndex = 2;
            this.label4.Text = "検索条件列";
            // 
            // columnValues
            // 
            this.columnValues.Location = new System.Drawing.Point(278, 3);
            this.columnValues.Name = "columnValues";
            this.columnValues.Size = new System.Drawing.Size(144, 20);
            this.columnValues.TabIndex = 3;
            // 
            // reportText
            // 
            this.reportText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.reportText.Location = new System.Drawing.Point(3, 257);
            this.reportText.Multiline = true;
            this.reportText.Name = "reportText";
            this.reportText.Size = new System.Drawing.Size(488, 108);
            this.reportText.TabIndex = 3;
            // 
            // flowLayoutPanel3
            // 
            this.flowLayoutPanel3.Controls.Add(this.label3);
            this.flowLayoutPanel3.Controls.Add(this.searchValues);
            this.flowLayoutPanel3.Location = new System.Drawing.Point(3, 141);
            this.flowLayoutPanel3.Name = "flowLayoutPanel3";
            this.flowLayoutPanel3.Size = new System.Drawing.Size(480, 100);
            this.flowLayoutPanel3.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "検索する値";
            // 
            // searchValues
            // 
            this.searchValues.Location = new System.Drawing.Point(76, 3);
            this.searchValues.Multiline = true;
            this.searchValues.Name = "searchValues";
            this.searchValues.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.searchValues.Size = new System.Drawing.Size(196, 97);
            this.searchValues.TabIndex = 1;
            // 
            // statusBarControl
            // 
            this.statusBarControl.ImageScalingSize = new System.Drawing.Size(17, 17);
            this.statusBarControl.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusText});
            this.statusBarControl.Location = new System.Drawing.Point(0, 376);
            this.statusBarControl.Name = "statusBarControl";
            this.statusBarControl.Size = new System.Drawing.Size(501, 22);
            this.statusBarControl.TabIndex = 1;
            this.statusBarControl.Text = "statusStrip1";
            // 
            // statusText
            // 
            this.statusText.Name = "statusText";
            this.statusText.Size = new System.Drawing.Size(167, 17);
            this.statusText.Text = "Excelファイルを選択してください";
            // 
            // flowLayoutPanel4
            // 
            this.flowLayoutPanel4.Controls.Add(this.label1);
            this.flowLayoutPanel4.Controls.Add(this.outputDirectory);
            this.flowLayoutPanel4.Controls.Add(this.browseOutputDirectoryButton);
            this.flowLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel4.Location = new System.Drawing.Point(3, 52);
            this.flowLayoutPanel4.Name = "flowLayoutPanel4";
            this.flowLayoutPanel4.Size = new System.Drawing.Size(488, 32);
            this.flowLayoutPanel4.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "出力先";
            // 
            // outputDirectory
            // 
            this.outputDirectory.Location = new System.Drawing.Point(55, 3);
            this.outputDirectory.Name = "outputDirectory";
            this.outputDirectory.Size = new System.Drawing.Size(260, 20);
            this.outputDirectory.TabIndex = 1;
            // 
            // browseOutputDirectoryButton
            // 
            this.browseOutputDirectoryButton.Location = new System.Drawing.Point(321, 3);
            this.browseOutputDirectoryButton.Name = "browseOutputDirectoryButton";
            this.browseOutputDirectoryButton.Size = new System.Drawing.Size(28, 23);
            this.browseOutputDirectoryButton.TabIndex = 2;
            this.browseOutputDirectoryButton.Text = "...";
            this.browseOutputDirectoryButton.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(501, 398);
            this.Controls.Add(this.statusBarControl);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Form1";
            this.Text = "EffectiveEx";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel2.PerformLayout();
            this.flowLayoutPanel3.ResumeLayout(false);
            this.flowLayoutPanel3.PerformLayout();
            this.statusBarControl.ResumeLayout(false);
            this.statusBarControl.PerformLayout();
            this.flowLayoutPanel4.ResumeLayout(false);
            this.flowLayoutPanel4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button openButton;
        private System.Windows.Forms.Button deleteRowButton;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox sheetNameCombo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox columnValues;
        private System.Windows.Forms.TextBox reportText;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox searchValues;
        private System.Windows.Forms.StatusStrip statusBarControl;
        private System.Windows.Forms.ToolStripStatusLabel statusText;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox outputDirectory;
        private System.Windows.Forms.Button browseOutputDirectoryButton;
    }
}

