using System;
using System.Data;
using System.Data.OleDb;


namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!" + DateTime.Now.ToString("yyyy/MM/dd"));
            Console.WriteLine("Hello World!" + DateTime.Now.ToString("yyyy/MM/dd"));

        }
    }
    class Access
    {
        OleDbConnection oledb = new OleDbConnection(@"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = D:\OneDrive\ProjectTest\Database3.mdb");
        
        public Access()
            {
            oledb.Open();
             }
        public void Get()
        {
            string sql = "select * from 表1";
            //獲取表1的內容
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oledb); //建立適配物件
            DataTable dt = new DataTable(); //新建表物件
            dbDataAdapter.Fill(dt); //用適配物件填充表物件
            foreach (DataRow item in dt.Rows)
            {
                Console.WriteLine(item[0] + " | " + item[1]);
            }
        }

        public void Find()
        {
            string sql = "select * from 表1 WHERE 暱稱='LanQ'";
            //獲取表1中暱稱為LanQ的內容
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oledb); //建立適配物件
            DataTable dt = new DataTable(); //新建表物件
            dbDataAdapter.Fill(dt); //用適配物件填充表物件
            foreach (DataRow item in dt.Rows)
            {
                Console.WriteLine(item[0] + " | " + item[1]);
            }
        }
        public bool Add()
        {
            string sql = "insert into 表1 (暱稱,賬號) values ('LanQ','2545493686')";
            //往表1新增一條記錄，暱稱是LanQ，賬號是2545493686
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oledb);
            int i = oleDbCommand.ExecuteNonQuery(); //返回被修改的數目
            return i > 0;
        }
        public bool Del()
        {
            string sql = "delete from 表1 where 暱稱='LANQ'";
            //刪除暱稱為LanQ的記錄
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oledb);
            int i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }
        public bool Change()
        {
            string sql = "update 表1 set 賬號='233333' where 暱稱='東熊'";
            //將表1中暱稱為東熊的賬號修改成233333
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oledb);
            int i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }
    }
}
