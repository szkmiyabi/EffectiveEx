﻿namespace EffectiveEx
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.openButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.sheetNameCombo = new System.Windows.Forms.ComboBox();
            this.worksheetPreviewButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.label4 = new System.Windows.Forms.Label();
            this.columnValues = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.skipRowNumbers = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.withoutConditionColNum = new System.Windows.Forms.NumericUpDown();
            this.reportText = new System.Windows.Forms.TextBox();
            this.flowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.searchValues = new System.Windows.Forms.TextBox();
            this.searchValResultWithoutCheck = new System.Windows.Forms.CheckBox();
            this.searchValsResultButton = new System.Windows.Forms.Button();
            this.flowLayoutPanel5 = new System.Windows.Forms.FlowLayoutPanel();
            this.reportClearButton = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.flowLayoutPanel4 = new System.Windows.Forms.FlowLayoutPanel();
            this.deleteRowButton = new System.Windows.Forms.Button();
            this.lpReportFormatButton = new System.Windows.Forms.Button();
            this.mergedCellAttentionButton = new System.Windows.Forms.Button();
            this.BookFooterClearButton = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.flowLayoutPanel9 = new System.Windows.Forms.FlowLayoutPanel();
            this.ankQueryButton = new System.Windows.Forms.Button();
            this.cvEncodePointButton = new System.Windows.Forms.Button();
            this.cvQueryOutputColumnButton = new System.Windows.Forms.Button();
            this.statusBarControl = new System.Windows.Forms.StatusStrip();
            this.statusText = new System.Windows.Forms.ToolStripStatusLabel();
            this.savePDFButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.columnValues)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.skipRowNumbers)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.withoutConditionColNum)).BeginInit();
            this.flowLayoutPanel3.SuspendLayout();
            this.flowLayoutPanel5.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.flowLayoutPanel4.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.flowLayoutPanel9.SuspendLayout();
            this.statusBarControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.reportText, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel3, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel5, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.tabControl1, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 6;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.25373F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 124F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.74627F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 103F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 43F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 177F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(608, 558);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.openButton);
            this.flowLayoutPanel1.Controls.Add(this.label2);
            this.flowLayoutPanel1.Controls.Add(this.sheetNameCombo);
            this.flowLayoutPanel1.Controls.Add(this.worksheetPreviewButton);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(602, 48);
            this.flowLayoutPanel1.TabIndex = 0;
            // 
            // openButton
            // 
            this.openButton.Location = new System.Drawing.Point(3, 3);
            this.openButton.Name = "openButton";
            this.openButton.Size = new System.Drawing.Size(112, 25);
            this.openButton.TabIndex = 1;
            this.openButton.Text = "開く";
            this.openButton.UseVisualStyleBackColor = true;
            this.openButton.Click += new System.EventHandler(this.openButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(121, 10);
            this.label2.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "対象シート";
            // 
            // sheetNameCombo
            // 
            this.sheetNameCombo.FormattingEnabled = true;
            this.sheetNameCombo.Location = new System.Drawing.Point(199, 6);
            this.sheetNameCombo.Margin = new System.Windows.Forms.Padding(3, 6, 3, 3);
            this.sheetNameCombo.Name = "sheetNameCombo";
            this.sheetNameCombo.Size = new System.Drawing.Size(166, 22);
            this.sheetNameCombo.TabIndex = 3;
            // 
            // worksheetPreviewButton
            // 
            this.worksheetPreviewButton.Location = new System.Drawing.Point(371, 3);
            this.worksheetPreviewButton.Name = "worksheetPreviewButton";
            this.worksheetPreviewButton.Size = new System.Drawing.Size(97, 25);
            this.worksheetPreviewButton.TabIndex = 5;
            this.worksheetPreviewButton.Text = "プレビュー";
            this.worksheetPreviewButton.UseVisualStyleBackColor = true;
            this.worksheetPreviewButton.Click += new System.EventHandler(this.worksheetPreviewButton_Click);
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Controls.Add(this.label4);
            this.flowLayoutPanel2.Controls.Add(this.columnValues);
            this.flowLayoutPanel2.Controls.Add(this.label5);
            this.flowLayoutPanel2.Controls.Add(this.skipRowNumbers);
            this.flowLayoutPanel2.Controls.Add(this.label6);
            this.flowLayoutPanel2.Controls.Add(this.withoutConditionColNum);
            this.flowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 181);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(602, 50);
            this.flowLayoutPanel2.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 6);
            this.label4.Margin = new System.Windows.Forms.Padding(3, 6, 3, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(112, 15);
            this.label4.TabIndex = 2;
            this.label4.Text = "条件判定列番号";
            // 
            // columnValues
            // 
            this.columnValues.Location = new System.Drawing.Point(121, 3);
            this.columnValues.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.columnValues.Name = "columnValues";
            this.columnValues.Size = new System.Drawing.Size(48, 21);
            this.columnValues.TabIndex = 5;
            this.columnValues.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(175, 6);
            this.label5.Margin = new System.Windows.Forms.Padding(3, 6, 3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(110, 15);
            this.label5.TabIndex = 2;
            this.label5.Text = "処理スキップ行数";
            // 
            // skipRowNumbers
            // 
            this.skipRowNumbers.Location = new System.Drawing.Point(291, 3);
            this.skipRowNumbers.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.skipRowNumbers.Name = "skipRowNumbers";
            this.skipRowNumbers.Size = new System.Drawing.Size(52, 21);
            this.skipRowNumbers.TabIndex = 4;
            this.skipRowNumbers.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(349, 6);
            this.label6.Margin = new System.Windows.Forms.Padding(3, 6, 3, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(112, 15);
            this.label6.TabIndex = 0;
            this.label6.Text = "絞込判定列番号";
            // 
            // withoutConditionColNum
            // 
            this.withoutConditionColNum.Location = new System.Drawing.Point(467, 3);
            this.withoutConditionColNum.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.withoutConditionColNum.Name = "withoutConditionColNum";
            this.withoutConditionColNum.Size = new System.Drawing.Size(52, 21);
            this.withoutConditionColNum.TabIndex = 2;
            this.withoutConditionColNum.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // reportText
            // 
            this.reportText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.reportText.Location = new System.Drawing.Point(3, 383);
            this.reportText.Multiline = true;
            this.reportText.Name = "reportText";
            this.reportText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.reportText.Size = new System.Drawing.Size(602, 172);
            this.reportText.TabIndex = 3;
            // 
            // flowLayoutPanel3
            // 
            this.flowLayoutPanel3.Controls.Add(this.label3);
            this.flowLayoutPanel3.Controls.Add(this.searchValues);
            this.flowLayoutPanel3.Controls.Add(this.searchValResultWithoutCheck);
            this.flowLayoutPanel3.Controls.Add(this.searchValsResultButton);
            this.flowLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel3.Location = new System.Drawing.Point(3, 237);
            this.flowLayoutPanel3.Name = "flowLayoutPanel3";
            this.flowLayoutPanel3.Size = new System.Drawing.Size(602, 97);
            this.flowLayoutPanel3.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 6);
            this.label3.Margin = new System.Windows.Forms.Padding(3, 6, 3, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 15);
            this.label3.TabIndex = 0;
            this.label3.Text = "検索値";
            // 
            // searchValues
            // 
            this.searchValues.Location = new System.Drawing.Point(61, 3);
            this.searchValues.Multiline = true;
            this.searchValues.Name = "searchValues";
            this.searchValues.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.searchValues.Size = new System.Drawing.Size(238, 84);
            this.searchValues.TabIndex = 1;
            // 
            // searchValResultWithoutCheck
            // 
            this.searchValResultWithoutCheck.AutoSize = true;
            this.searchValResultWithoutCheck.Location = new System.Drawing.Point(305, 9);
            this.searchValResultWithoutCheck.Margin = new System.Windows.Forms.Padding(3, 9, 3, 3);
            this.searchValResultWithoutCheck.Name = "searchValResultWithoutCheck";
            this.searchValResultWithoutCheck.Size = new System.Drawing.Size(113, 19);
            this.searchValResultWithoutCheck.TabIndex = 1;
            this.searchValResultWithoutCheck.Text = "検索値で絞込";
            this.searchValResultWithoutCheck.UseVisualStyleBackColor = true;
            // 
            // searchValsResultButton
            // 
            this.searchValsResultButton.Location = new System.Drawing.Point(424, 3);
            this.searchValsResultButton.Name = "searchValsResultButton";
            this.searchValsResultButton.Size = new System.Drawing.Size(108, 25);
            this.searchValsResultButton.TabIndex = 3;
            this.searchValsResultButton.Text = "値検索集計";
            this.searchValsResultButton.UseVisualStyleBackColor = true;
            this.searchValsResultButton.Click += new System.EventHandler(this.searchValsResultButton_Click);
            // 
            // flowLayoutPanel5
            // 
            this.flowLayoutPanel5.Controls.Add(this.reportClearButton);
            this.flowLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Right;
            this.flowLayoutPanel5.Location = new System.Drawing.Point(511, 340);
            this.flowLayoutPanel5.Name = "flowLayoutPanel5";
            this.flowLayoutPanel5.Size = new System.Drawing.Size(94, 37);
            this.flowLayoutPanel5.TabIndex = 5;
            // 
            // reportClearButton
            // 
            this.reportClearButton.Location = new System.Drawing.Point(3, 3);
            this.reportClearButton.Name = "reportClearButton";
            this.reportClearButton.Size = new System.Drawing.Size(75, 28);
            this.reportClearButton.TabIndex = 0;
            this.reportClearButton.Text = "クリア";
            this.reportClearButton.UseVisualStyleBackColor = true;
            this.reportClearButton.Click += new System.EventHandler(this.reportClearButton_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(10, 64);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(10);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(588, 104);
            this.tabControl1.TabIndex = 6;
            // 
            // tabPage1
            // 
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tabPage1.Controls.Add(this.flowLayoutPanel4);
            this.tabPage1.Location = new System.Drawing.Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(580, 76);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "MAIN";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // flowLayoutPanel4
            // 
            this.flowLayoutPanel4.Controls.Add(this.deleteRowButton);
            this.flowLayoutPanel4.Controls.Add(this.lpReportFormatButton);
            this.flowLayoutPanel4.Controls.Add(this.mergedCellAttentionButton);
            this.flowLayoutPanel4.Controls.Add(this.BookFooterClearButton);
            this.flowLayoutPanel4.Controls.Add(this.savePDFButton);
            this.flowLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel4.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel4.Name = "flowLayoutPanel4";
            this.flowLayoutPanel4.Size = new System.Drawing.Size(572, 68);
            this.flowLayoutPanel4.TabIndex = 0;
            // 
            // deleteRowButton
            // 
            this.deleteRowButton.Location = new System.Drawing.Point(3, 3);
            this.deleteRowButton.Name = "deleteRowButton";
            this.deleteRowButton.Size = new System.Drawing.Size(102, 25);
            this.deleteRowButton.TabIndex = 2;
            this.deleteRowButton.Text = "行削除";
            this.deleteRowButton.UseVisualStyleBackColor = true;
            this.deleteRowButton.Click += new System.EventHandler(this.deleteRowButton_Click);
            // 
            // lpReportFormatButton
            // 
            this.lpReportFormatButton.Location = new System.Drawing.Point(111, 3);
            this.lpReportFormatButton.Name = "lpReportFormatButton";
            this.lpReportFormatButton.Size = new System.Drawing.Size(98, 25);
            this.lpReportFormatButton.TabIndex = 4;
            this.lpReportFormatButton.Text = "LPR表整形";
            this.lpReportFormatButton.UseVisualStyleBackColor = true;
            this.lpReportFormatButton.Click += new System.EventHandler(this.lpReportFormatButton_Click);
            // 
            // mergedCellAttentionButton
            // 
            this.mergedCellAttentionButton.Location = new System.Drawing.Point(215, 3);
            this.mergedCellAttentionButton.Name = "mergedCellAttentionButton";
            this.mergedCellAttentionButton.Size = new System.Drawing.Size(112, 25);
            this.mergedCellAttentionButton.TabIndex = 5;
            this.mergedCellAttentionButton.Text = "結合セル強調";
            this.mergedCellAttentionButton.UseVisualStyleBackColor = true;
            this.mergedCellAttentionButton.Click += new System.EventHandler(this.mergedCellAttentionButton_Click);
            // 
            // BookFooterClearButton
            // 
            this.BookFooterClearButton.Location = new System.Drawing.Point(333, 3);
            this.BookFooterClearButton.Name = "BookFooterClearButton";
            this.BookFooterClearButton.Size = new System.Drawing.Size(100, 25);
            this.BookFooterClearButton.TabIndex = 6;
            this.BookFooterClearButton.Text = "フッター消去";
            this.BookFooterClearButton.UseVisualStyleBackColor = true;
            this.BookFooterClearButton.Click += new System.EventHandler(this.BookFooterClearButton_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tabPage2.Controls.Add(this.flowLayoutPanel9);
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(580, 76);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "SUB";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // flowLayoutPanel9
            // 
            this.flowLayoutPanel9.Controls.Add(this.ankQueryButton);
            this.flowLayoutPanel9.Controls.Add(this.cvEncodePointButton);
            this.flowLayoutPanel9.Controls.Add(this.cvQueryOutputColumnButton);
            this.flowLayoutPanel9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel9.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel9.Name = "flowLayoutPanel9";
            this.flowLayoutPanel9.Size = new System.Drawing.Size(572, 68);
            this.flowLayoutPanel9.TabIndex = 0;
            // 
            // ankQueryButton
            // 
            this.ankQueryButton.Location = new System.Drawing.Point(3, 3);
            this.ankQueryButton.Name = "ankQueryButton";
            this.ankQueryButton.Size = new System.Drawing.Size(79, 25);
            this.ankQueryButton.TabIndex = 6;
            this.ankQueryButton.Text = "ANK抽出";
            this.ankQueryButton.UseVisualStyleBackColor = true;
            this.ankQueryButton.Click += new System.EventHandler(this.ankQueryButton_Click);
            // 
            // cvEncodePointButton
            // 
            this.cvEncodePointButton.Location = new System.Drawing.Point(88, 3);
            this.cvEncodePointButton.Name = "cvEncodePointButton";
            this.cvEncodePointButton.Size = new System.Drawing.Size(96, 23);
            this.cvEncodePointButton.TabIndex = 7;
            this.cvEncodePointButton.Text = "CV有無点付";
            this.cvEncodePointButton.UseVisualStyleBackColor = true;
            this.cvEncodePointButton.Click += new System.EventHandler(this.cvEncodePointButton_Click);
            // 
            // cvQueryOutputColumnButton
            // 
            this.cvQueryOutputColumnButton.Location = new System.Drawing.Point(190, 3);
            this.cvQueryOutputColumnButton.Name = "cvQueryOutputColumnButton";
            this.cvQueryOutputColumnButton.Size = new System.Drawing.Size(115, 23);
            this.cvQueryOutputColumnButton.TabIndex = 8;
            this.cvQueryOutputColumnButton.Text = "CV有無列出力";
            this.cvQueryOutputColumnButton.UseVisualStyleBackColor = true;
            this.cvQueryOutputColumnButton.Click += new System.EventHandler(this.cvQueryOutputColumnButton_Click);
            // 
            // statusBarControl
            // 
            this.statusBarControl.ImageScalingSize = new System.Drawing.Size(17, 17);
            this.statusBarControl.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusText});
            this.statusBarControl.Location = new System.Drawing.Point(0, 563);
            this.statusBarControl.Name = "statusBarControl";
            this.statusBarControl.Size = new System.Drawing.Size(608, 26);
            this.statusBarControl.TabIndex = 1;
            this.statusBarControl.Text = "statusStrip1";
            // 
            // statusText
            // 
            this.statusText.Name = "statusText";
            this.statusText.Size = new System.Drawing.Size(190, 20);
            this.statusText.Text = "Excelファイルを選択してください";
            // 
            // savePDFButton
            // 
            this.savePDFButton.Location = new System.Drawing.Point(439, 3);
            this.savePDFButton.Name = "savePDFButton";
            this.savePDFButton.Size = new System.Drawing.Size(102, 25);
            this.savePDFButton.TabIndex = 7;
            this.savePDFButton.Text = "PDF保存";
            this.savePDFButton.UseVisualStyleBackColor = true;
            this.savePDFButton.Click += new System.EventHandler(this.savePDFButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(608, 589);
            this.Controls.Add(this.statusBarControl);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "EffectiveEx";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.columnValues)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.skipRowNumbers)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.withoutConditionColNum)).EndInit();
            this.flowLayoutPanel3.ResumeLayout(false);
            this.flowLayoutPanel3.PerformLayout();
            this.flowLayoutPanel5.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.flowLayoutPanel4.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.flowLayoutPanel9.ResumeLayout(false);
            this.statusBarControl.ResumeLayout(false);
            this.statusBarControl.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button openButton;
        private System.Windows.Forms.Button deleteRowButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox sheetNameCombo;
        private System.Windows.Forms.TextBox reportText;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox searchValues;
        private System.Windows.Forms.StatusStrip statusBarControl;
        private System.Windows.Forms.ToolStripStatusLabel statusText;
        private System.Windows.Forms.Button searchValsResultButton;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel5;
        private System.Windows.Forms.Button reportClearButton;
        private System.Windows.Forms.CheckBox searchValResultWithoutCheck;
        private System.Windows.Forms.Button lpReportFormatButton;
        private System.Windows.Forms.Button worksheetPreviewButton;
        private System.Windows.Forms.Button mergedCellAttentionButton;
        private System.Windows.Forms.Button ankQueryButton;
        private System.Windows.Forms.Button cvEncodePointButton;
        private System.Windows.Forms.Button cvQueryOutputColumnButton;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel4;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel9;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown columnValues;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown skipRowNumbers;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown withoutConditionColNum;
        private System.Windows.Forms.Button BookFooterClearButton;
        private System.Windows.Forms.Button savePDFButton;
    }
}

