namespace OutlookSenderStatistics
{
    partial class FormMain
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
            toolStripContainer1 = new ToolStripContainer();
            statusStrip1 = new StatusStrip();
            toolStripStatusLabel1 = new ToolStripStatusLabel();
            toolStripProgressBar1 = new ToolStripProgressBar();
            tableLayoutPanel1 = new TableLayoutPanel();
            label1 = new Label();
            flowLayoutPanelDelete = new FlowLayoutPanel();
            labelDeleteFromTheseSenders = new Label();
            checkBoxPartialMatchAddress = new CheckBox();
            textBoxAddressesToDelete = new TextBox();
            labelAndSenderAgeOlderThan = new Label();
            numericUpDownMinAgeToDelete = new NumericUpDown();
            checkBoxMinSizeToDelete = new CheckBox();
            numericUpDownMinSizeToDelete = new NumericUpDown();
            dataGridViewEmailFolders = new DataGridView();
            bindingSourceFolders = new BindingSource(components);
            labelDelete = new Label();
            flowLayoutPanelCount = new FlowLayoutPanel();
            buttonStartCounting = new Button();
            buttonStopCounting = new Button();
            flowLayoutPanel1 = new FlowLayoutPanel();
            buttonStartDeletion = new Button();
            buttonStopDeletion = new Button();
            toolStrip1 = new ToolStrip();
            toolStripButtonOpenOutlook = new ToolStripButton();
            saveFileDialog1 = new SaveFileDialog();
            isSelectedDataGridViewCheckBoxColumn = new DataGridViewCheckBoxColumn();
            Folder = new DataGridViewTextBoxColumn();
            toolStripContainer1.BottomToolStripPanel.SuspendLayout();
            toolStripContainer1.ContentPanel.SuspendLayout();
            toolStripContainer1.TopToolStripPanel.SuspendLayout();
            toolStripContainer1.SuspendLayout();
            statusStrip1.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            flowLayoutPanelDelete.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)numericUpDownMinAgeToDelete).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownMinSizeToDelete).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridViewEmailFolders).BeginInit();
            ((System.ComponentModel.ISupportInitialize)bindingSourceFolders).BeginInit();
            flowLayoutPanelCount.SuspendLayout();
            flowLayoutPanel1.SuspendLayout();
            toolStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // toolStripContainer1
            // 
            // 
            // toolStripContainer1.BottomToolStripPanel
            // 
            toolStripContainer1.BottomToolStripPanel.Controls.Add(statusStrip1);
            // 
            // toolStripContainer1.ContentPanel
            // 
            toolStripContainer1.ContentPanel.Controls.Add(tableLayoutPanel1);
            toolStripContainer1.ContentPanel.Size = new Size(1502, 800);
            toolStripContainer1.Dock = DockStyle.Fill;
            toolStripContainer1.Location = new Point(0, 0);
            toolStripContainer1.Name = "toolStripContainer1";
            toolStripContainer1.Size = new Size(1502, 866);
            toolStripContainer1.TabIndex = 0;
            toolStripContainer1.Text = "toolStripContainer1";
            // 
            // toolStripContainer1.TopToolStripPanel
            // 
            toolStripContainer1.TopToolStripPanel.Controls.Add(toolStrip1);
            // 
            // statusStrip1
            // 
            statusStrip1.Dock = DockStyle.None;
            statusStrip1.ImageScalingSize = new Size(24, 24);
            statusStrip1.Items.AddRange(new ToolStripItem[] { toolStripStatusLabel1, toolStripProgressBar1 });
            statusStrip1.Location = new Point(0, 0);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(1502, 32);
            statusStrip1.TabIndex = 0;
            // 
            // toolStripStatusLabel1
            // 
            toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            toolStripStatusLabel1.Size = new Size(60, 25);
            toolStripStatusLabel1.Text = "Ready";
            // 
            // toolStripProgressBar1
            // 
            toolStripProgressBar1.Name = "toolStripProgressBar1";
            toolStripProgressBar1.Size = new Size(100, 24);
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle());
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle());
            tableLayoutPanel1.Controls.Add(label1, 0, 0);
            tableLayoutPanel1.Controls.Add(flowLayoutPanelDelete, 1, 1);
            tableLayoutPanel1.Controls.Add(dataGridViewEmailFolders, 0, 1);
            tableLayoutPanel1.Controls.Add(labelDelete, 1, 0);
            tableLayoutPanel1.Controls.Add(flowLayoutPanelCount, 0, 2);
            tableLayoutPanel1.Controls.Add(flowLayoutPanel1, 1, 2);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 3;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 10F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 80F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 10F));
            tableLayoutPanel1.Size = new Size(1502, 800);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(211, 25);
            label1.TabIndex = 0;
            label1.Text = "Choose Inboxes to count";
            // 
            // flowLayoutPanelDelete
            // 
            flowLayoutPanelDelete.Controls.Add(labelDeleteFromTheseSenders);
            flowLayoutPanelDelete.Controls.Add(checkBoxPartialMatchAddress);
            flowLayoutPanelDelete.Controls.Add(textBoxAddressesToDelete);
            flowLayoutPanelDelete.Controls.Add(labelAndSenderAgeOlderThan);
            flowLayoutPanelDelete.Controls.Add(numericUpDownMinAgeToDelete);
            flowLayoutPanelDelete.Controls.Add(checkBoxMinSizeToDelete);
            flowLayoutPanelDelete.Controls.Add(numericUpDownMinSizeToDelete);
            flowLayoutPanelDelete.Dock = DockStyle.Fill;
            flowLayoutPanelDelete.FlowDirection = FlowDirection.TopDown;
            flowLayoutPanelDelete.Location = new Point(754, 83);
            flowLayoutPanelDelete.Name = "flowLayoutPanelDelete";
            flowLayoutPanelDelete.Size = new Size(745, 634);
            flowLayoutPanelDelete.TabIndex = 3;
            // 
            // labelDeleteFromTheseSenders
            // 
            labelDeleteFromTheseSenders.AutoSize = true;
            labelDeleteFromTheseSenders.Location = new Point(3, 0);
            labelDeleteFromTheseSenders.Name = "labelDeleteFromTheseSenders";
            labelDeleteFromTheseSenders.Size = new Size(389, 25);
            labelDeleteFromTheseSenders.TabIndex = 0;
            labelDeleteFromTheseSenders.Text = "From these senders (one email address per line)";
            // 
            // checkBoxPartialMatchAddress
            // 
            checkBoxPartialMatchAddress.AutoSize = true;
            checkBoxPartialMatchAddress.Location = new Point(3, 28);
            checkBoxPartialMatchAddress.Name = "checkBoxPartialMatchAddress";
            checkBoxPartialMatchAddress.Size = new Size(187, 29);
            checkBoxPartialMatchAddress.TabIndex = 1;
            checkBoxPartialMatchAddress.Text = "partial match is OK";
            checkBoxPartialMatchAddress.UseVisualStyleBackColor = true;
            // 
            // textBoxAddressesToDelete
            // 
            textBoxAddressesToDelete.Location = new Point(3, 63);
            textBoxAddressesToDelete.Multiline = true;
            textBoxAddressesToDelete.Name = "textBoxAddressesToDelete";
            textBoxAddressesToDelete.Size = new Size(487, 259);
            textBoxAddressesToDelete.TabIndex = 1;
            // 
            // labelAndSenderAgeOlderThan
            // 
            labelAndSenderAgeOlderThan.AutoSize = true;
            labelAndSenderAgeOlderThan.Location = new Point(3, 325);
            labelAndSenderAgeOlderThan.Name = "labelAndSenderAgeOlderThan";
            labelAndSenderAgeOlderThan.Size = new Size(217, 25);
            labelAndSenderAgeOlderThan.TabIndex = 2;
            labelAndSenderAgeOlderThan.Text = "And age older than (days)\r\n";
            // 
            // numericUpDownMinAgeToDelete
            // 
            numericUpDownMinAgeToDelete.Location = new Point(3, 353);
            numericUpDownMinAgeToDelete.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            numericUpDownMinAgeToDelete.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numericUpDownMinAgeToDelete.Name = "numericUpDownMinAgeToDelete";
            numericUpDownMinAgeToDelete.Size = new Size(180, 31);
            numericUpDownMinAgeToDelete.TabIndex = 7;
            numericUpDownMinAgeToDelete.Value = new decimal(new int[] { 30, 0, 0, 0 });
            // 
            // checkBoxMinSizeToDelete
            // 
            checkBoxMinSizeToDelete.AutoSize = true;
            checkBoxMinSizeToDelete.Location = new Point(3, 390);
            checkBoxMinSizeToDelete.Name = "checkBoxMinSizeToDelete";
            checkBoxMinSizeToDelete.Size = new Size(275, 29);
            checkBoxMinSizeToDelete.TabIndex = 5;
            checkBoxMinSizeToDelete.Text = "And mail size larger than (mb)";
            checkBoxMinSizeToDelete.UseVisualStyleBackColor = true;
            // 
            // numericUpDownMinSizeToDelete
            // 
            numericUpDownMinSizeToDelete.Location = new Point(3, 425);
            numericUpDownMinSizeToDelete.Name = "numericUpDownMinSizeToDelete";
            numericUpDownMinSizeToDelete.Size = new Size(180, 31);
            numericUpDownMinSizeToDelete.TabIndex = 8;
            numericUpDownMinSizeToDelete.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // dataGridViewEmailFolders
            // 
            dataGridViewEmailFolders.AllowUserToAddRows = false;
            dataGridViewEmailFolders.AllowUserToDeleteRows = false;
            dataGridViewEmailFolders.AutoGenerateColumns = false;
            dataGridViewEmailFolders.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewEmailFolders.Columns.AddRange(new DataGridViewColumn[] { isSelectedDataGridViewCheckBoxColumn, Folder });
            dataGridViewEmailFolders.DataSource = bindingSourceFolders;
            dataGridViewEmailFolders.Dock = DockStyle.Fill;
            dataGridViewEmailFolders.Location = new Point(3, 83);
            dataGridViewEmailFolders.Name = "dataGridViewEmailFolders";
            dataGridViewEmailFolders.RowHeadersWidth = 62;
            dataGridViewEmailFolders.Size = new Size(745, 634);
            dataGridViewEmailFolders.TabIndex = 1;
            // 
            // bindingSourceFolders
            // 
            bindingSourceFolders.AllowNew = false;
            bindingSourceFolders.DataSource = typeof(InboxSelection);
            // 
            // labelDelete
            // 
            labelDelete.AutoSize = true;
            labelDelete.Location = new Point(754, 0);
            labelDelete.Name = "labelDelete";
            labelDelete.Size = new Size(62, 25);
            labelDelete.TabIndex = 2;
            labelDelete.Text = "Delete";
            // 
            // flowLayoutPanelCount
            // 
            flowLayoutPanelCount.Controls.Add(buttonStartCounting);
            flowLayoutPanelCount.Controls.Add(buttonStopCounting);
            flowLayoutPanelCount.Dock = DockStyle.Fill;
            flowLayoutPanelCount.Location = new Point(3, 723);
            flowLayoutPanelCount.Name = "flowLayoutPanelCount";
            flowLayoutPanelCount.Size = new Size(745, 74);
            flowLayoutPanelCount.TabIndex = 4;
            // 
            // buttonStartCounting
            // 
            buttonStartCounting.Location = new Point(3, 3);
            buttonStartCounting.Name = "buttonStartCounting";
            buttonStartCounting.Size = new Size(112, 34);
            buttonStartCounting.TabIndex = 0;
            buttonStartCounting.Text = "Start counting";
            buttonStartCounting.UseVisualStyleBackColor = true;
            buttonStartCounting.Click += buttonStartCounting_Click;
            // 
            // buttonStopCounting
            // 
            buttonStopCounting.Location = new Point(121, 3);
            buttonStopCounting.Name = "buttonStopCounting";
            buttonStopCounting.Size = new Size(112, 34);
            buttonStopCounting.TabIndex = 1;
            buttonStopCounting.Text = "Stop";
            buttonStopCounting.UseVisualStyleBackColor = true;
            buttonStopCounting.Click += buttonStopCounting_Click;
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Controls.Add(buttonStartDeletion);
            flowLayoutPanel1.Controls.Add(buttonStopDeletion);
            flowLayoutPanel1.Dock = DockStyle.Fill;
            flowLayoutPanel1.Location = new Point(754, 723);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(745, 74);
            flowLayoutPanel1.TabIndex = 5;
            // 
            // buttonStartDeletion
            // 
            buttonStartDeletion.Location = new Point(3, 3);
            buttonStartDeletion.Name = "buttonStartDeletion";
            buttonStartDeletion.Size = new Size(112, 34);
            buttonStartDeletion.TabIndex = 0;
            buttonStartDeletion.Text = "Start";
            buttonStartDeletion.UseVisualStyleBackColor = true;
            buttonStartDeletion.Click += buttonStartDeletion_Click;
            // 
            // buttonStopDeletion
            // 
            buttonStopDeletion.Location = new Point(121, 3);
            buttonStopDeletion.Name = "buttonStopDeletion";
            buttonStopDeletion.Size = new Size(112, 34);
            buttonStopDeletion.TabIndex = 1;
            buttonStopDeletion.Text = "Stop";
            buttonStopDeletion.UseVisualStyleBackColor = true;
            buttonStopDeletion.Click += buttonStopDeletion_Click;
            // 
            // toolStrip1
            // 
            toolStrip1.Dock = DockStyle.None;
            toolStrip1.ImageScalingSize = new Size(24, 24);
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripButtonOpenOutlook });
            toolStrip1.Location = new Point(4, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(148, 34);
            toolStrip1.TabIndex = 0;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonOpenOutlook
            // 
            toolStripButtonOpenOutlook.DisplayStyle = ToolStripItemDisplayStyle.Text;
            toolStripButtonOpenOutlook.ImageTransparentColor = Color.Magenta;
            toolStripButtonOpenOutlook.Name = "toolStripButtonOpenOutlook";
            toolStripButtonOpenOutlook.Size = new Size(130, 29);
            toolStripButtonOpenOutlook.Text = "Open Outlook";
            toolStripButtonOpenOutlook.Click += toolStripButtonAttachToOutlook_Click;
            // 
            // saveFileDialog1
            // 
            saveFileDialog1.DefaultExt = "csv";
            // 
            // isSelectedDataGridViewCheckBoxColumn
            // 
            isSelectedDataGridViewCheckBoxColumn.DataPropertyName = "IsSelected";
            isSelectedDataGridViewCheckBoxColumn.HeaderText = "Include in counting";
            isSelectedDataGridViewCheckBoxColumn.MinimumWidth = 8;
            isSelectedDataGridViewCheckBoxColumn.Name = "isSelectedDataGridViewCheckBoxColumn";
            isSelectedDataGridViewCheckBoxColumn.Width = 200;
            // 
            // Folder
            // 
            Folder.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Folder.DataPropertyName = "Folder";
            Folder.HeaderText = "Folder";
            Folder.MinimumWidth = 8;
            Folder.Name = "Folder";
            // 
            // FormMain
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1502, 866);
            Controls.Add(toolStripContainer1);
            Name = "FormMain";
            Text = "Outlook Sender Stastics";
            FormClosing += FormMain_FormClosing;
            Load += FormMain_Load;
            toolStripContainer1.BottomToolStripPanel.ResumeLayout(false);
            toolStripContainer1.BottomToolStripPanel.PerformLayout();
            toolStripContainer1.ContentPanel.ResumeLayout(false);
            toolStripContainer1.TopToolStripPanel.ResumeLayout(false);
            toolStripContainer1.TopToolStripPanel.PerformLayout();
            toolStripContainer1.ResumeLayout(false);
            toolStripContainer1.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.PerformLayout();
            flowLayoutPanelDelete.ResumeLayout(false);
            flowLayoutPanelDelete.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)numericUpDownMinAgeToDelete).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDownMinSizeToDelete).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridViewEmailFolders).EndInit();
            ((System.ComponentModel.ISupportInitialize)bindingSourceFolders).EndInit();
            flowLayoutPanelCount.ResumeLayout(false);
            flowLayoutPanel1.ResumeLayout(false);
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private ToolStripContainer toolStripContainer1;
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButtonOpenOutlook;
        private TableLayoutPanel tableLayoutPanel1;
        private Label label1;
        private DataGridView dataGridViewEmailFolders;
        private BindingSource bindingSourceFolders;
        private SaveFileDialog saveFileDialog1;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private ToolStripProgressBar toolStripProgressBar1;
        private Label labelDelete;
        private FlowLayoutPanel flowLayoutPanelDelete;
        private Label labelDeleteFromTheseSenders;
        private CheckBox checkBoxPartialMatchAddress;
        private TextBox textBoxAddressesToDelete;
        private Label labelAndSenderAgeOlderThan;
        private FlowLayoutPanel flowLayoutPanelCount;
        private Button buttonStartCounting;
        private Button buttonStopCounting;
        private FlowLayoutPanel flowLayoutPanel1;
        private Button buttonStartDeletion;
        private Button buttonStopDeletion;
        private CheckBox checkBoxMinSizeToDelete;
        private NumericUpDown numericUpDownMinAgeToDelete;
        private NumericUpDown numericUpDownMinSizeToDelete;
        private DataGridViewCheckBoxColumn isSelectedDataGridViewCheckBoxColumn;
        private DataGridViewTextBoxColumn Folder;
    }
}
