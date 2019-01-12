namespace Word2Access
{
    partial class Form1
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.TabPage = new System.Windows.Forms.TabControl();
            this.ManagerPage = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tablenametips = new System.Windows.Forms.Label();
            this.CreatetableButton = new System.Windows.Forms.Button();
            this.Colums = new System.Windows.Forms.TextBox();
            this.columtips = new System.Windows.Forms.Label();
            this.TableName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.TableDetail = new System.Windows.Forms.GroupBox();
            this.MatchListBox = new System.Windows.Forms.ListBox();
            this.Addmatch = new System.Windows.Forms.Button();
            this.MatchContent = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.ColumBox = new System.Windows.Forms.ListBox();
            this.SelectTablename = new System.Windows.Forms.Label();
            this.TableManagerBox = new System.Windows.Forms.GroupBox();
            this.TableList = new System.Windows.Forms.ListView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.ImportWordBox = new System.Windows.Forms.GroupBox();
            this.ImportLog = new System.Windows.Forms.TextBox();
            this.WordFilePath = new System.Windows.Forms.TextBox();
            this.ChooseWordFileButton = new System.Windows.Forms.Button();
            this.SerachPage = new System.Windows.Forms.TabPage();
            this.ImportButton = new System.Windows.Forms.Button();
            this.Tips = new System.Windows.Forms.Label();
            this.TabPage.SuspendLayout();
            this.ManagerPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.TableDetail.SuspendLayout();
            this.TableManagerBox.SuspendLayout();
            this.ImportWordBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabPage
            // 
            this.TabPage.Controls.Add(this.ManagerPage);
            this.TabPage.Controls.Add(this.SerachPage);
            this.TabPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TabPage.Location = new System.Drawing.Point(0, 0);
            this.TabPage.Name = "TabPage";
            this.TabPage.SelectedIndex = 0;
            this.TabPage.Size = new System.Drawing.Size(1388, 657);
            this.TabPage.TabIndex = 0;
            // 
            // ManagerPage
            // 
            this.ManagerPage.Controls.Add(this.groupBox1);
            this.ManagerPage.Controls.Add(this.TableDetail);
            this.ManagerPage.Controls.Add(this.TableManagerBox);
            this.ManagerPage.Controls.Add(this.ImportWordBox);
            this.ManagerPage.Location = new System.Drawing.Point(4, 24);
            this.ManagerPage.Name = "ManagerPage";
            this.ManagerPage.Padding = new System.Windows.Forms.Padding(3);
            this.ManagerPage.Size = new System.Drawing.Size(1380, 629);
            this.ManagerPage.TabIndex = 0;
            this.ManagerPage.Text = "Manager";
            this.ManagerPage.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.LightGray;
            this.groupBox1.Controls.Add(this.tablenametips);
            this.groupBox1.Controls.Add(this.CreatetableButton);
            this.groupBox1.Controls.Add(this.Colums);
            this.groupBox1.Controls.Add(this.columtips);
            this.groupBox1.Controls.Add(this.TableName);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(9, 7);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(538, 247);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "CreateTable";
            // 
            // tablenametips
            // 
            this.tablenametips.AutoSize = true;
            this.tablenametips.Location = new System.Drawing.Point(264, 34);
            this.tablenametips.Name = "tablenametips";
            this.tablenametips.Size = new System.Drawing.Size(175, 14);
            this.tablenametips.TabIndex = 5;
            this.tablenametips.Text = "（不要有空格和特殊字符）";
            // 
            // CreatetableButton
            // 
            this.CreatetableButton.Location = new System.Drawing.Point(223, 205);
            this.CreatetableButton.Name = "CreatetableButton";
            this.CreatetableButton.Size = new System.Drawing.Size(87, 27);
            this.CreatetableButton.TabIndex = 4;
            this.CreatetableButton.Text = "创建数据表";
            this.CreatetableButton.UseVisualStyleBackColor = true;
            this.CreatetableButton.Click += new System.EventHandler(this.CreatetableButton_Click);
            // 
            // Colums
            // 
            this.Colums.Location = new System.Drawing.Point(64, 86);
            this.Colums.Multiline = true;
            this.Colums.Name = "Colums";
            this.Colums.Size = new System.Drawing.Size(466, 111);
            this.Colums.TabIndex = 3;
            // 
            // columtips
            // 
            this.columtips.AutoSize = true;
            this.columtips.Location = new System.Drawing.Point(7, 69);
            this.columtips.Name = "columtips";
            this.columtips.Size = new System.Drawing.Size(252, 14);
            this.columtips.TabIndex = 2;
            this.columtips.Text = "字段（半角逗号隔开,不要有特殊字符）";
            // 
            // TableName
            // 
            this.TableName.Location = new System.Drawing.Point(64, 24);
            this.TableName.Name = "TableName";
            this.TableName.Size = new System.Drawing.Size(175, 23);
            this.TableName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 14);
            this.label1.TabIndex = 0;
            this.label1.Text = "表名";
            // 
            // TableDetail
            // 
            this.TableDetail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableDetail.BackColor = System.Drawing.Color.BlanchedAlmond;
            this.TableDetail.Controls.Add(this.MatchListBox);
            this.TableDetail.Controls.Add(this.Addmatch);
            this.TableDetail.Controls.Add(this.MatchContent);
            this.TableDetail.Controls.Add(this.label2);
            this.TableDetail.Controls.Add(this.ColumBox);
            this.TableDetail.Controls.Add(this.SelectTablename);
            this.TableDetail.Location = new System.Drawing.Point(561, 7);
            this.TableDetail.Name = "TableDetail";
            this.TableDetail.Size = new System.Drawing.Size(813, 357);
            this.TableDetail.TabIndex = 2;
            this.TableDetail.TabStop = false;
            this.TableDetail.Text = "TableDetail";
            // 
            // MatchListBox
            // 
            this.MatchListBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.MatchListBox.FormattingEnabled = true;
            this.MatchListBox.ItemHeight = 20;
            this.MatchListBox.Location = new System.Drawing.Point(191, 69);
            this.MatchListBox.Name = "MatchListBox";
            this.MatchListBox.Size = new System.Drawing.Size(616, 184);
            this.MatchListBox.TabIndex = 6;
            this.MatchListBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.MatchListBox_DrawItem);
            this.MatchListBox.SelectedIndexChanged += new System.EventHandler(this.MatchListBox_SelectedIndexChanged);
            // 
            // Addmatch
            // 
            this.Addmatch.Location = new System.Drawing.Point(734, 270);
            this.Addmatch.Name = "Addmatch";
            this.Addmatch.Size = new System.Drawing.Size(73, 69);
            this.Addmatch.TabIndex = 5;
            this.Addmatch.Text = "新增匹配";
            this.Addmatch.UseVisualStyleBackColor = true;
            this.Addmatch.Click += new System.EventHandler(this.Savematch_Click);
            // 
            // MatchContent
            // 
            this.MatchContent.Location = new System.Drawing.Point(191, 270);
            this.MatchContent.Multiline = true;
            this.MatchContent.Name = "MatchContent";
            this.MatchContent.Size = new System.Drawing.Size(537, 69);
            this.MatchContent.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(189, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(329, 14);
            this.label2.TabIndex = 2;
            this.label2.Text = "匹配模式(一行一个匹配方法，回车换行，逗号隔开)";
            // 
            // ColumBox
            // 
            this.ColumBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.ColumBox.FormattingEnabled = true;
            this.ColumBox.ItemHeight = 20;
            this.ColumBox.Location = new System.Drawing.Point(22, 69);
            this.ColumBox.Name = "ColumBox";
            this.ColumBox.Size = new System.Drawing.Size(151, 264);
            this.ColumBox.TabIndex = 1;
            this.ColumBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ColumBox_DrawItem);
            // 
            // SelectTablename
            // 
            this.SelectTablename.AutoSize = true;
            this.SelectTablename.Location = new System.Drawing.Point(20, 34);
            this.SelectTablename.Name = "SelectTablename";
            this.SelectTablename.Size = new System.Drawing.Size(35, 14);
            this.SelectTablename.TabIndex = 0;
            this.SelectTablename.Text = "字段";
            // 
            // TableManagerBox
            // 
            this.TableManagerBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.TableManagerBox.Controls.Add(this.TableList);
            this.TableManagerBox.Location = new System.Drawing.Point(1, 261);
            this.TableManagerBox.Name = "TableManagerBox";
            this.TableManagerBox.Size = new System.Drawing.Size(553, 356);
            this.TableManagerBox.TabIndex = 1;
            this.TableManagerBox.TabStop = false;
            this.TableManagerBox.Text = "TableManagerBox";
            // 
            // TableList
            // 
            this.TableList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableList.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.TableList.LargeImageList = this.imageList1;
            this.TableList.Location = new System.Drawing.Point(7, 23);
            this.TableList.Name = "TableList";
            this.TableList.Size = new System.Drawing.Size(538, 325);
            this.TableList.TabIndex = 0;
            this.TableList.UseCompatibleStateImageBehavior = false;
            this.TableList.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.TableList_ItemSelectionChanged);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "logo.png");
            // 
            // ImportWordBox
            // 
            this.ImportWordBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ImportWordBox.BackColor = System.Drawing.Color.LightSkyBlue;
            this.ImportWordBox.Controls.Add(this.Tips);
            this.ImportWordBox.Controls.Add(this.ImportButton);
            this.ImportWordBox.Controls.Add(this.ImportLog);
            this.ImportWordBox.Controls.Add(this.WordFilePath);
            this.ImportWordBox.Controls.Add(this.ChooseWordFileButton);
            this.ImportWordBox.Location = new System.Drawing.Point(561, 370);
            this.ImportWordBox.Name = "ImportWordBox";
            this.ImportWordBox.Size = new System.Drawing.Size(808, 245);
            this.ImportWordBox.TabIndex = 0;
            this.ImportWordBox.TabStop = false;
            this.ImportWordBox.Text = "ImportWordBox";
            // 
            // ImportLog
            // 
            this.ImportLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ImportLog.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.ImportLog.ForeColor = System.Drawing.SystemColors.Window;
            this.ImportLog.Location = new System.Drawing.Point(29, 105);
            this.ImportLog.Multiline = true;
            this.ImportLog.Name = "ImportLog";
            this.ImportLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.ImportLog.Size = new System.Drawing.Size(743, 134);
            this.ImportLog.TabIndex = 2;
            // 
            // WordFilePath
            // 
            this.WordFilePath.Location = new System.Drawing.Point(29, 74);
            this.WordFilePath.Name = "WordFilePath";
            this.WordFilePath.Size = new System.Drawing.Size(534, 23);
            this.WordFilePath.TabIndex = 1;
            // 
            // ChooseWordFileButton
            // 
            this.ChooseWordFileButton.Location = new System.Drawing.Point(583, 70);
            this.ChooseWordFileButton.Name = "ChooseWordFileButton";
            this.ChooseWordFileButton.Size = new System.Drawing.Size(87, 27);
            this.ChooseWordFileButton.TabIndex = 0;
            this.ChooseWordFileButton.Text = "浏览";
            this.ChooseWordFileButton.UseVisualStyleBackColor = true;
            this.ChooseWordFileButton.Click += new System.EventHandler(this.ChooseWordFileButton_Click);
            // 
            // SerachPage
            // 
            this.SerachPage.Location = new System.Drawing.Point(4, 24);
            this.SerachPage.Name = "SerachPage";
            this.SerachPage.Padding = new System.Windows.Forms.Padding(3);
            this.SerachPage.Size = new System.Drawing.Size(1380, 629);
            this.SerachPage.TabIndex = 1;
            this.SerachPage.Text = "Search";
            this.SerachPage.UseVisualStyleBackColor = true;
            // 
            // ImportButton
            // 
            this.ImportButton.Location = new System.Drawing.Point(685, 72);
            this.ImportButton.Name = "ImportButton";
            this.ImportButton.Size = new System.Drawing.Size(87, 27);
            this.ImportButton.TabIndex = 3;
            this.ImportButton.Text = "导入";
            this.ImportButton.UseVisualStyleBackColor = true;
            this.ImportButton.Click += new System.EventHandler(this.ImportButton_Click);
            // 
            // Tips
            // 
            this.Tips.AutoSize = true;
            this.Tips.Location = new System.Drawing.Point(29, 35);
            this.Tips.Name = "Tips";
            this.Tips.Size = new System.Drawing.Size(161, 14);
            this.Tips.TabIndex = 4;
            this.Tips.Text = "Tips :  当前匹配模式为";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1388, 657);
            this.Controls.Add(this.TabPage);
            this.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Name = "Form1";
            this.Text = "Form1";
            this.TabPage.ResumeLayout(false);
            this.ManagerPage.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.TableDetail.ResumeLayout(false);
            this.TableDetail.PerformLayout();
            this.TableManagerBox.ResumeLayout(false);
            this.ImportWordBox.ResumeLayout(false);
            this.ImportWordBox.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl TabPage;
        private System.Windows.Forms.TabPage ManagerPage;
        private System.Windows.Forms.GroupBox ImportWordBox;
        private System.Windows.Forms.TextBox WordFilePath;
        private System.Windows.Forms.Button ChooseWordFileButton;
        private System.Windows.Forms.TabPage SerachPage;
        private System.Windows.Forms.TextBox ImportLog;
        private System.Windows.Forms.GroupBox TableDetail;
        private System.Windows.Forms.GroupBox TableManagerBox;
        private System.Windows.Forms.ListView TableList;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button CreatetableButton;
        private System.Windows.Forms.TextBox Colums;
        private System.Windows.Forms.Label columtips;
        private System.Windows.Forms.TextBox TableName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label tablenametips;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label SelectTablename;
        private System.Windows.Forms.ListBox ColumBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox MatchContent;
        private System.Windows.Forms.Button Addmatch;
        private System.Windows.Forms.ListBox MatchListBox;
        private System.Windows.Forms.Button ImportButton;
        private System.Windows.Forms.Label Tips;
    }
}

