using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using ADOX;
using System.IO;
namespace Word2Access
{
    public partial class Form1 : Form
    {
        Access MyDataSet;
        public Form1()
        {
            InitializeComponent();
            initDataSet();
            freshTableList();
            freshTableTree();
        }


        private void initDataSet()
        {
            string path = System.Environment.CurrentDirectory + "\\All.mdb";
            bool flag = AccessDbHelper.CreateAccessDb(path);
            if (flag) Console.WriteLine("create success");
            else Console.WriteLine("already exits");

            MyDataSet = new Access();
            TabPage.SelectedIndex = 1;
        }




        private void CreatetableButton_Click(object sender, EventArgs e)
        {
            String mytablename = TableName.Text.Trim();
            String columnames = Colums.Text.Trim();

            if (string.IsNullOrWhiteSpace(mytablename))
            {
                tablenametips.Text = "表名不能为空";
                tablenametips.ForeColor = Color.Red;
                return;
            }

            if (MyDataSet.GetTableNameList().Contains(mytablename))
            {
                tablenametips.Text = "表名已存在，请更换";
                tablenametips.ForeColor = Color.Red;
                return;
            }

            if (string.IsNullOrWhiteSpace(columnames))
            {
                tablenametips.Text = "字段名不能为空";
                tablenametips.ForeColor = Color.Red;
                return;
            }

            freshCreateTable();

            string[] mycolums = columnames.Split(',');

            string s = "表名为：" + mytablename + "\r\n包含字段有： ";
            for (int i = 0; i < mycolums.Length; i++) s += (i + 1) + " : " + mycolums[i] + ",    ";
            s += "是否确认建表？";
            DialogResult result = MessageBox.Show(s, "确认健表", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result == DialogResult.OK)
            {
                ADOX.Column[] columns = new ADOX.Column[mycolums.Length + 2];
                columns[0] = new ADOX.Column() { Name = "id", Type = DataTypeEnum.adInteger, DefinedSize = 9 };

                for (int i = 0; i < mycolums.Length; i++)
                {
                    columns[i + 1] = new ADOX.Column() { Name = mycolums[i], Type = DataTypeEnum.adLongVarWChar, DefinedSize = 500 };
                }
                columns[mycolums.Length + 1] = new ADOX.Column() { Name = "importfilename", Type = DataTypeEnum.adLongVarWChar, DefinedSize = 500 };
                bool flag = AccessDbHelper.CreateAccessTable(System.Environment.CurrentDirectory + "\\All.mdb", mytablename, columns);
                if (flag)
                {
                    MessageBox.Show("建表成功");
                    CreatemMatchFile(mytablename);
                    freshTableList();
                    freshTableTree();
                }
                else MessageBox.Show("建表失败");

            }


        }

        private void freshCreateTable()
        {
            tablenametips.Text = "（不要有空格和特殊字符）";
            tablenametips.ForeColor = Color.Black;
            tablenametips.Text = "字段（半角逗号隔开,不要有特殊字符）";
            tablenametips.ForeColor = Color.Black;
            return;
        }

        private void freshTableList()
        {
            TableList.Items.Clear();
            MyDataSet.Reconnect();
            TableList.BeginUpdate();
            List<string> tablenamelist = MyDataSet.GetTableNameList();
            for (int i = 0; i < tablenamelist.Count; i++)
            {
                ListViewItem item = new ListViewItem();
                item.Text = tablenamelist[i];
                item.ImageIndex = 0;
                TableList.Items.Add(item);
            }
            TableList.EndUpdate();
        }

        private void TableList_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (TableList.SelectedItems.Count >= 1)
            {
                ColumBox.Items.Clear();
                TableDetail.Text = TableList.SelectedItems[0].Text;
                CreatemMatchFile(TableDetail.Text);
                String[] cloumnames = MyDataSet.GetTableColumn(TableList.SelectedItems[0].Text);
                for (int i = 0; i < cloumnames.Length; i++)
                {
                    if (cloumnames[i] != "id" && cloumnames[i] != "importfilename")
                        ColumBox.Items.Add(cloumnames[i]);
                }
                readMatchFile(TableDetail.Text);
            }

        }



        private void CreatemMatchFile(string tablename)
        {
            string path = System.Environment.CurrentDirectory + "\\match\\" + tablename + ".match";
            if (!File.Exists(path))
            {   //参数1：要创建的文件路径，包含文件名称、后缀等
                FileStream fs = File.Create(path);
                fs.Close();
            }

        }

        private void readMatchFile(string tablename)
        {
            MatchListBox.Items.Clear();
            string path = System.Environment.CurrentDirectory + "\\match\\" + tablename + ".match";
            StreamReader sr = new StreamReader(path, Encoding.UTF8);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                MatchListBox.Items.Add(line);
            }
            if (MatchListBox.Items.Count > 0) MatchListBox.SelectedIndex = 0;
            else Tips.Text = "Tips :  无匹配模式，请先添加匹配";
            sr.Close();
        }

        private void Savematch_Click(object sender, EventArgs e)
        {
            string path = System.Environment.CurrentDirectory + "\\match\\" + TableDetail.Text + ".match";
            FileStream fs = new FileStream(path, FileMode.Append);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            if (MatchListBox.Items.Contains(MatchContent.Text))
            {
                MessageBox.Show("匹配已存在");
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();
                return;
            }
            bool flag = checkmatchline(MatchContent.Text);
            if (flag)
            {
                sw.WriteLine(MatchContent.Text);
                MatchListBox.Items.Add(MatchContent.Text);
                MessageBox.Show("保存成功");
            }
            else
            {
                MessageBox.Show("格式错误");
            }

            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();

        }

        private bool checkmatchline(string text)
        {
            string[] mymatch = text.Split(',');
            if (mymatch.Length != ColumBox.Items.Count) return false;
            if (IsSameWithHashSet(mymatch)) return false;
            return true;
        }

        static bool IsSameWithHashSet(string[] arr)
        {
            ISet<string> set = new HashSet<string>();

            for (var i = 0; i < arr.Length; i++)
            {
                // 这里可利用该元素来实现统计重复的原理有哪些，及重复个数。
                //bool state = set.Add(arr[i]); // 如果返回false，表示set中已经有该元素。
                //Console.WriteLine(state);
                set.Add(arr[i]);
            }

            return set.Count != arr.Length;
        }

        private void ChooseWordFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件";
            dialog.Filter = "Word文件(*.doc,*.docx)|*.doc;*.docx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                WordFilePath.Text = dialog.FileName;
            }
        }

        private void ImportWordFile()
        {
            ImportLog.Clear();
            String filename = WordFilePath.Text;
            if (string.IsNullOrWhiteSpace(filename))
            {
                MessageBox.Show("请先选择导入文件");
                return;
            }

            //载入Word文档
            Document document = new Document(filename);

            Section section = document.Sections[0];
            int i = 1;
            foreach (Spire.Doc.Table table in section.Tables)
            {
                ImportLog.AppendText("导入表格" + i + "，自动查找匹配模式开始...\r\n");
                string[] cls = getDocColumsFromDocTable(table);
                int matchindex = getMatchIndex(cls);
                if (matchindex < 0)
                {
                    ImportLog.AppendText("查找匹配模式失败，请添加合适的匹配模式\r\n");
                }
                else
                {
                    ImportLog.AppendText("查找匹配模式成功，匹配模式为：" + MatchListBox.Items[matchindex].ToString() + "\r\n");

                    ImportTableData(matchindex, table);
                }
                i++;
            }


            //Spire.Doc.Table table = section.Tables[0] as Spire.Doc.Table;
            ////实例化StringBuilder类
            //StringBuilder sb = new StringBuilder();

            ////遍历表格中的段落并提取文本
            //foreach (TableRow row in table.Rows)
            //{
            //    foreach (TableCell cell in row.Cells)
            //    {
            //        foreach (Paragraph paragraph in cell.Paragraphs)
            //        {
            //            sb.AppendLine(paragraph.Text);
            //        }
            //    }
            //}
            //ImportLog.Text = sb.ToString();
        }

        private void ImportTableData(int matchindex, Spire.Doc.Table table)
        {
            ImportLog.AppendText("开始导入数据\r\n");
            string strInsertHead = getinsertsqlhead(matchindex);

            string[] insertsql = new string[ColumBox.Items.Count];
            for (int i = 0; i < ColumBox.Items.Count - 1; i++)
            {
                insertsql[i] = "''";
            }
            string[] cls = getDocColumsFromDocTable(table);
            List<int> AccesscolumsIndex = getAccesscolumsIndex(matchindex, cls);

            string shortFileName = WordFilePath.Text.Substring(WordFilePath.Text.LastIndexOf('\\') + 1);
            for (int i = 1; i < table.Rows.Count; i++)
            {
                string strInsert = strInsertHead;

                for (int j = 0; j < AccesscolumsIndex.Count; j++)
                {
                    if (AccesscolumsIndex[j] >= 0)
                    {
                        string cellcontent = "";
                        foreach (Paragraph paragraph in table.Rows[i].Cells[j].Paragraphs)
                        {
                            cellcontent += paragraph.Text + " ";
                        }

                        insertsql[AccesscolumsIndex[j]] = "'" + cellcontent.Trim() + "'";
                    }
                }
                for (int k = 0; k < ColumBox.Items.Count; k++)
                {
                    strInsert += insertsql[k] + ",";
                }

                strInsert += "'" + shortFileName + "' )";

                bool flag = MyDataSet.Add(strInsert);
                if (flag) ImportLog.AppendText("导入第" + i + "条数据成功," + strInsert + "\r\n");
                else
                {
                    ImportLog.AppendText("导入第" + i + "条数据失败," + strInsert + "\r\n");

                }
            }
            ImportLog.AppendText("导入数据完成\r\n");
            freshTableTree();

        }

        private List<int> getAccesscolumsIndex(int matchindex, String[] cls)
        {
            List<int> indexs = new List<int>();
            List<string> vv = MatchListBox.Items[matchindex].ToString().Split(',').ToList();
            for (int k = 0; k < cls.Length; k++)
            {
                if (vv.Contains(cls[k])) indexs.Add(vv.IndexOf(cls[k]));
                else indexs.Add(-1);
            }
            return indexs;
        }

        private string getinsertsqlhead(int matchindex)
        {
            string s = " INSERT INTO " + TableDetail.Text;// + " ( bookid , booktitle , bookauthor , bookprice , bookstock ) VALUES ( ";
            s += " ( ";
            int i = 0;
            for (i = 0; i < ColumBox.Items.Count; i++)
            {
                s += ColumBox.Items[i].ToString() + " , ";
            }
            s += "importfilename ) VALUES(";

            return s;
        }

        private int getMatchIndex(string[] cls)
        {
            int maxMatch = 0, mathindex = -1;
            for (int i = 0; i < MatchListBox.Items.Count; i++)
            {
                int tmpMatch = ComputeMatch(cls, MatchListBox.Items[i].ToString());
                if (tmpMatch > maxMatch)
                {
                    maxMatch = tmpMatch;
                    mathindex = i;
                }
            }
            return mathindex;
        }

        private int ComputeMatch(string[] cls, string v)
        {
            int matchsize = 0;
            string[] vv = v.Split(',');

            for (int i = 0; i < cls.Length; i++)
            {
                if (vv.Contains(cls[i]))
                {
                    matchsize++;
                }

            }
            int emptysize = 0;
            for (int j = 0; j < vv.Length; j++)
            {
                if (vv[j] == "")
                    emptysize++;
            }
            if ((emptysize + matchsize) == vv.Length) return matchsize;
            else return 0;
        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            ImportWordFile();
        }

        private void ColumBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            e.DrawFocusRectangle();
            e.Graphics.DrawString(ColumBox.Items[e.Index].ToString(), e.Font, new SolidBrush(Color.Black), e.Bounds);

        }

        private void MatchListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            e.DrawFocusRectangle();
            e.Graphics.DrawString(MatchListBox.Items[e.Index].ToString(), e.Font, new SolidBrush(Color.Black), e.Bounds);
        }

        private void MatchListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Tips.Text = "Tips :  当前匹配模式为： " + MatchListBox.SelectedItem.ToString();
        }

        private string[] getDocColumsFromDocTable(Spire.Doc.Table table)
        {

            //遍历表格中的段落并提取文本
            string[] cls = new string[table.Rows[0].Cells.Count];
            for (int i = 0; i < table.Rows[0].Cells.Count; i++)
            {
                cls[i] = table.Rows[0].Cells[i].Paragraphs[0].Text;

            }
            return cls;
        }


        private void freshTableTree()
        {
            TableTree.Nodes.Clear();
            List<string> tablenamelist = MyDataSet.GetTableNameList();
            TreeNode rootnode = new TreeNode();
            rootnode.Text = "所有数据表";
            for (int i = 0; i < tablenamelist.Count; i++)
            {
                TreeNode tablenode = new TreeNode();
                tablenode.Text = tablenamelist[i];
                List<string> importfilelist = MyDataSet.GetImportFileList(tablenamelist[i]);
                for (int j = 0; j < importfilelist.Count; j++)
                {
                    TreeNode importfilenode = new TreeNode();
                    importfilenode.Text = importfilelist[j];
                    tablenode.Nodes.Add(importfilenode);
                }
                rootnode.Nodes.Add(tablenode);
            }
            TableTree.Nodes.Add(rootnode);
            rootnode.ExpandAll();
        }

        private void TableTree_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode selectnode = e.Node;
            string tablename = "", importfilename = "";
            if (selectnode.Level == 0) return;
            else if(selectnode.Level == 1)
            {
                tablename = selectnode.Text;
                DataTable searchdata = MyDataSet.getSearchData(tablename);
                dataTableToListview(searchdata, ResultList);
            }
            else if(selectnode.Level == 2)
            {
                importfilename = selectnode.Text;
                tablename = selectnode.Parent.Text;
                DataTable searchdata = MyDataSet.getSearchData(tablename, importfilename);
                dataTableToListview(searchdata, ResultList);
            }
            freshSearchTableLayOutPanel(tablename);
            
            
        }

        private void freshSearchTableLayOutPanel(string tablename)
        {
            string[] colums = MyDataSet.GetTableColumn(tablename);
            tableLayoutPanel2.Controls.Clear();
            tableLayoutPanel2.ColumnCount = 1;
            tableLayoutPanel2.RowCount = 0;
            for (int i = 0;i< colums.Length ; i++)
            {
                if (colums[i] != "id" && colums[i] != "importfilename")
                {
                    tableLayoutPanel2.RowCount = tableLayoutPanel2.RowCount = 2;
                    Label label1 = new Label();
                    label1.Text = colums[i];
                    label1.Width = 250;
                    label1.TextAlign = ContentAlignment.BottomLeft;
                    tableLayoutPanel2.Controls.Add(label1);
                    System.Windows.Forms.TextBox tb = new System.Windows.Forms.TextBox();
                    tb.Width = tableLayoutPanel2.Width - 10;
                    tableLayoutPanel2.Controls.Add(tb);
                }

            }

        }

        /// <summary>
        /// 为ListView绑定DataTable数据项
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="lv">ListView控件</param>
        public void dataTableToListview(DataTable dt, ListView lv)
        {
            if (dt != null)
            {
                lv.View = System.Windows.Forms.View.Details;
                lv.GridLines = true;//显示网格线
                lv.Items.Clear();//所有的项
                lv.Columns.Clear();//标题
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    lv.Columns.Add(dt.Columns[i].Caption.ToString());//增加标题
                }
                this.label3.Text = "总共查询到：" + dt.Rows.Count + " 条记录。";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ListViewItem lvi = new ListViewItem((i+1)+"");
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        // lvi.ImageIndex = 0;
                        if(dt.Columns[j].Caption.ToString() == "字段4")
                        {
                            string[] s = dt.Rows[i][j].ToString().Split(' ');
                            string ss = "";
                            int index = 1;
                            for (int k = 0; k < s.Length; k++) {
                                if (!string.IsNullOrEmpty(s[k]))
                                {
                                    ss += s[k] + "(" + index + ") ";
                                    index++;
                                }
                            }
                            lvi.SubItems.Add(ss);
                        }
                        else
                            lvi.SubItems.Add(dt.Rows[i][j].ToString());
                    }
                    lv.Items.Add(lvi);
                }
                lv.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);//调整列的宽度
            }
        }

        static public void listViewToDataTable(ListView lv, DataTable dt)
        {
            int i, j;
            DataRow dr;
            dt.Clear();
            dt.Columns.Clear();
            //生成DataTable列头
            for (i = 0; i < lv.Columns.Count; i++)
            {
                dt.Columns.Add(lv.Columns[i].Text.Trim(), typeof(String));
            }
            //每行内容
            for (i = 0; i < lv.Items.Count; i++)
            {
                dr = dt.NewRow();
                for (j = 0; j < lv.Columns.Count; j++)
                {
                    dr[j] = lv.Items[i].SubItems[j].Text.Trim();
                }
                dt.Rows.Add(dr);
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            string sql = "select id, * from ";
            string searchsql = getSearchsql();
            string tablename = "", importfilename = "";
            if (string.IsNullOrEmpty(searchsql)) return;
            else {
                if (TableTree.SelectedNode.Level == 1)
                {
                    tablename = TableTree.SelectedNode.Text;
                    sql += tablename + " " + searchsql + " id > 0";
                    DataTable searchdata = MyDataSet.getSearchDataBydql(sql);
                    dataTableToListview(searchdata, ResultList);
                }
                else if (TableTree.SelectedNode.Level == 2)
                {
                    importfilename = TableTree.SelectedNode.Text;
                    tablename = TableTree.SelectedNode.Parent.Text;
                    sql += tablename + " " + searchsql + " importfilename = '"+ importfilename + "'";
                    DataTable searchdata = MyDataSet.getSearchDataBydql(sql);
                    dataTableToListview(searchdata, ResultList);
                }
                

            }

        }

        private string getSearchsql()
        {
            string sql = "where ";
            for(int i = 0; i< tableLayoutPanel2.Controls.Count; i = i+2)
            {
                Label lb = (Label)tableLayoutPanel2.Controls[i];
                System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)tableLayoutPanel2.Controls[i+1];
                if (tb.Text != "") {
                    sql += lb.Text + " LIKE '%" + tb.Text + "%' and ";
                }
            }
            if(sql == " where ") return null;
            return sql;
        }

        private void DeleteTableButton_Click(object sender, EventArgs e)
        {
            string tablename = TableDetail.Text;
            DialogResult result = MessageBox.Show("确认删除表"+tablename+"吗？", "确认删除吗", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result == DialogResult.OK)
            {
                bool flag = MyDataSet.DelTable(tablename);
                if (flag)
                {
                    MessageBox.Show("数据表" + tablename + "已删除");
                    ColumBox.Items.Clear();
                    MatchListBox.Items.Clear();
                    // 返回与指定虚拟路径相对应的物理路径即绝对路径
                    string path = System.Environment.CurrentDirectory + "\\match\\" + tablename + ".match";
                    // 删除该文件
                    System.IO.File.Delete(path);
                    freshTableList();
                    freshTableTree();
                }
                else
                {
                    MessageBox.Show("数据表" + tablename + "删除失败");
                }
            }

        }

        private void Export2Word_Click(object sender, EventArgs e)
        {
            if (ResultList.Items.Count <= 0)
            {
                MessageBox.Show("无导出数据");
                return;
            }
            else
            {
                saveWordFile.Filter = "Word files (*.docx)|*.docx";
                saveWordFile.RestoreDirectory = true;
                if (saveWordFile.ShowDialog() == DialogResult.OK)
                {
                    saveMyWordFile(saveWordFile.FileName);
                }
            }
        }

        private void saveMyWordFile(string filename)
        {
            DataTable dt = new DataTable();
            listViewToDataTable(ResultList, dt);


            Document document = new Document();
            Section section = document.AddSection();
            
            Spire.Doc.Table table = section.AddTable(true);

            if (dt != null)
            {
                table.ResetCells(dt.Rows.Count + 1, dt.Columns.Count);
                TableRow row = table.Rows[0];
                row.IsHeader = true;
                row.HeightType = TableRowHeightType.Exactly;
                row.RowFormat.BackColor = Color.Gray;

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Paragraph para = row.Cells[i].AddParagraph();
                    TextRange TR = para.AppendText(dt.Columns[i].Caption.ToString());
                    TR.CharacterFormat.FontName = "宋体";
                    TR.CharacterFormat.FontSize = 14;
                    TR.CharacterFormat.Bold = true;

                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    TableRow row1 = table.Rows[i + 1];
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        Paragraph para = row1.Cells[j].AddParagraph();
                        TextRange TR = para.AppendText(dt.Rows[i][j].ToString());
                        TR.CharacterFormat.FontName = "宋体";
                        TR.CharacterFormat.FontSize = 12;
                    }
                }

                document.SaveToFile(filename);

                System.Diagnostics.Process.Start(filename);
            }
        }
    }
}
