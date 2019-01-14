using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace Word2Access
{
    class Access
    {
        static string str = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +  System.Environment.CurrentDirectory + "\\All.mdb";
        OleDbConnection oleDb = new OleDbConnection(str);

        public Access() //构造函数
        {
            oleDb.Open();
        }
        public void  Reconnect()
        {
            if (oleDb.State == ConnectionState.Closed)
                oleDb.Open();
            else {
                oleDb.Close();
                oleDb.Open();
            }
               
        }


        /// <summary>
        /// 取所有表名
        /// </summary>
        /// <returns></returns>
        public List<string> GetTableNameList()
        {
            List<string> list = new List<string>();
            try
            {
                if (oleDb.State == ConnectionState.Closed)
                    oleDb.Open();
                DataTable dt = oleDb.GetSchema("Tables");
                foreach (DataRow row in dt.Rows)
                {
                    if (row[3].ToString() == "TABLE")
                        list.Add(row[2].ToString());
                }
                return list;
            }
            catch (Exception e)
            { throw e; }
            finally {
                //if (oleDb.State == ConnectionState.Open) oleDb.Close(); oleDb.Dispose();
            }
        }

        /// <summary>
        /// 返回某一表的所有字段名
        /// </summary>
        public string[] GetTableColumn(string varTableName)
        {
            DataTable dt = new DataTable();
            try
            {
                if (oleDb.State == ConnectionState.Closed)
                    oleDb.Open();
                dt = oleDb.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, varTableName, null });
                int n = dt.Rows.Count;
                string[] strTable = new string[n];
                int m = dt.Columns.IndexOf("COLUMN_NAME");
                for (int i = 0; i < n; i++)
                {
                    DataRow m_DataRow = dt.Rows[i];
                    strTable[i] = m_DataRow.ItemArray.GetValue(m).ToString();
                }
                return strTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //oleDb.Close();
            }
        }

        public void Get()
        {
            string sql = "select * from 表1";
            //获取表1的内容
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb); //创建适配对象
            DataTable dt = new DataTable(); //新建表对象
            dbDataAdapter.Fill(dt); //用适配对象填充表对象
            foreach (DataRow item in dt.Rows)
            {
                Console.WriteLine(item[0] + " | " + item[1]);
            }
        }

        public void Find()
        {
            string sql = "select * from 表1 WHERE 昵称='LanQ'";
            //获取表1中昵称为东熊的内容
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb); //创建适配对象
            DataTable dt = new DataTable(); //新建表对象
            dbDataAdapter.Fill(dt); //用适配对象填充表对象
            foreach (DataRow item in dt.Rows)
            {
                Console.WriteLine(item[0] + " | " + item[1]);
            }
        }

        public bool Add(string sql)
        {
            //往表1添加一条记录，昵称是LanQ，账号是2545493686
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery(); //返回被修改的数目
            return i > 0;
        }
        public bool Del()
        {
            string sql = "delete from 表1 where 昵称='LANQ'";
            //删除昵称为LanQ的记录
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }
        public bool Change()
        {
            string sql = "update 表1 set 账号='233333' where 昵称='东熊'";
            //将表1中昵称为东熊的账号修改成233333
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }

        internal List<string> GetImportFileList(string tablename)
        {
            List<string> importfilelist = new List<string>();
            //SELECT importfilename FROM 测试 group by importfilename;
            string sql = "SELECT importfilename FROM " + tablename + " group by importfilename";
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb);
            DataTable dt = new DataTable(); //新建表对象
            dbDataAdapter.Fill(dt); //用适配对象填充表对象
            foreach (DataRow item in dt.Rows)
            {
                importfilelist.Add(item[0].ToString());
            }
            return importfilelist;
        }

        internal DataTable getSearchData(string tablename, string importfilename)
        {
            string sql = "SELECT id, * FROM " + tablename + " where importfilename = '"+importfilename+"'";
            
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb);
            DataTable dt = new DataTable();
            dbDataAdapter.Fill(dt);
            return dt;
        }

        internal DataTable getSearchData(string tablename)
        {
            string sql = "SELECT id, * FROM " + tablename;

            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb);
            DataTable dt = new DataTable();
            dbDataAdapter.Fill(dt);
            return dt;
        }

        internal DataTable getSearchDataBydql(string sql)
        {
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb);
            DataTable dt = new DataTable();
            dbDataAdapter.Fill(dt);
            return dt;
        }

        internal bool DelTable(string tablename)
        {
            string sql = "drop table " + tablename;
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery();
            return i >= 0;
        }
    }
}
