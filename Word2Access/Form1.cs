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
        }

        private void initDataSet()
        {
            string path = System.Environment.CurrentDirectory + "\\All.mdb";
            bool flag = AccessDbHelper.CreateAccessDb(path);
            if (flag) Console.WriteLine("create success");
            else Console.WriteLine("already exits");

            MyDataSet = new Access();
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
            for (int i = 0; i < mycolums.Length; i++) s += (i+1) + " : " + mycolums[i] + ",    ";
            s += "是否确认建表？";
            DialogResult result = MessageBox.Show(s, "确认健表", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result == DialogResult.OK)
            {
                ADOX.Column[] columns = new ADOX.Column[mycolums.Length + 1];
                columns[0] = new ADOX.Column() { Name = "id", Type = DataTypeEnum.adInteger, DefinedSize = 9 };

                for (int i = 0; i < mycolums.Length; i++)
                {
                    columns[i + 1] = new ADOX.Column() { Name = mycolums[i], Type = DataTypeEnum.adLongVarWChar, DefinedSize = 500 };
                }

                bool flag = AccessDbHelper.CreateAccessTable(System.Environment.CurrentDirectory + "\\All.mdb", mytablename, columns);
                if (flag)
                {
                    MessageBox.Show("建表成功");
                    CreatemMatchFile(mytablename);
                    freshTableList();
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
            for(int i = 0;i< tablenamelist.Count; i++)
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
                for(int i=0;i< cloumnames.Length; i++)
                {
                    if (cloumnames[i] != "id")
                        ColumBox.Items.Add(cloumnames[i]);
                }
                readMatchFile(TableDetail.Text);
            }
            
        }
        
        private void CreatemMatchFile( string tablename)
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
            StreamReader sr = new StreamReader(path, Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                MatchListBox.Items.Add(line);
            }
            if (MatchListBox.Items.Count > 0) MatchListBox.SelectedIndex = 0;
            else Tips.Text = "Tips :  无匹配模式，请先添加匹配" ;
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
            if(string.IsNullOrWhiteSpace(filename)){
                MessageBox.Show("请先选择导入文件");
                return; 
            }

            //载入Word文档
            Document document = new Document(filename);

            Section section = document.Sections[0];
            int i = 1;
            foreach (Spire.Doc.Table table in section.Tables) {
                ImportLog.AppendText("导入表格"+ i +"，自动查找匹配模式开始...\r\n"); 
                string[] cls = getDocColumsFromDocTable(table);
                int matchindex = getMatchIndex(cls);
                if (matchindex < 0) {
                    ImportLog.AppendText("查找匹配模式失败，请添加合适的匹配模式\r\n");
                }
                else
                {
                    ImportLog.AppendText("查找匹配模式成功，匹配模式为："+ MatchListBox.Items[matchindex].ToString() + "\r\n");
                    
                    ImportTableData(matchindex,table);
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
                insertsql[i] =  "''";
            }
            string[] cls = getDocColumsFromDocTable(table);
            List<int> AccesscolumsIndex = getAccesscolumsIndex(matchindex, cls);
            

            for (int i =1; i < table.Rows.Count; i++)
            {
                string strInsert = strInsertHead;

                for (int j= 0;j< AccesscolumsIndex.Count;j++) 
                {
                    if (AccesscolumsIndex[j] >= 0)
                    {
                        string cellcontent = "";
                        foreach (Paragraph paragraph in table.Rows[i].Cells[j].Paragraphs)
                        {
                            cellcontent += paragraph.Text + " ";
                        }

                        insertsql[AccesscolumsIndex[j]] = "'" +cellcontent.Trim()+"'";
                    }                 
                 }
                for (int k = 0; k < ColumBox.Items.Count - 1;k++)
                {
                    strInsert += insertsql[k] + ",";
                }
                strInsert += insertsql[ColumBox.Items.Count - 1] + " )";

                bool flag = MyDataSet.Add(strInsert);
                if(flag)ImportLog.AppendText("导入第"+i+"条数据成功,"+strInsert+"\r\n");
                else ImportLog.AppendText("导入第" + i + "条数据失败," + strInsert + "\r\n");
            }

            

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
            for(i =0;i< ColumBox.Items.Count - 1 ; i++)
            {
                s += ColumBox.Items[i].ToString() + " , ";
            }
            s += ColumBox.Items[ColumBox.Items.Count - 1].ToString() + " ) VALUES(";
                       
            return s;
        }

        private int getMatchIndex(string[] cls)
        {
            int maxMatch = 0, mathindex = -1;
            for (int i = 0; i < MatchListBox.Items.Count; i++) {
                int tmpMatch = ComputeMatch(cls, MatchListBox.Items[i].ToString());
                if (tmpMatch > maxMatch) {
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
            
            for (int i = 0; i < cls.Length; i++) {
                if (vv.Contains(cls[i]))
                {
                    matchsize++;
                }
                   
            }
            int emptysize = 0;
            for(int j =0;j < vv.Length; j++)
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
            for (int i = 0;i< table.Rows[0].Cells.Count;i++)
            {
                cls[i] = table.Rows[0].Cells[i].Paragraphs[0].Text;
                
            }
            return cls;
        }
    }
}
