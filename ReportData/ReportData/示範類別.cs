using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;

namespace ReportData
{
    class 連線類別
    {

        private readonly string 連線字串 = string.Empty;
        private SqlConnection 物件連線 = new SqlConnection();
        private SqlCommand 資料庫命令 = new SqlCommand();
        private SqlDataAdapter 資料表集合 = new SqlDataAdapter();
        public 連線類別()
        {
            try
            {
                連線字串 = ConfigurationManager.ConnectionStrings["ConnString"].ToString();  
            }
            catch (Exception 例外訊息)
            {
                //寫入Log紀錄
                throw new Exception(例外訊息.Message);
            }

        }
        private void 打開連線()
        {
            物件連線.ConnectionString = 連線字串;
            try
            {
                物件連線.Open();
            }
            catch (Exception 例外訊息)
            {
                //資料庫連接失敗,寫入Log紀錄
                throw new Exception(例外訊息.Message);
            }
        }
        private void 關閉連線()
        {
            物件連線.Close();
        }
        public IDataReader 取得SQL集合(string 指令, params SqlParameter[] 參數)
        {
            SqlDataReader 資料集合 = null;
            try
            {
                打開連線();
                資料庫命令 = new SqlCommand(指令, 物件連線);
                if (參數 != null && 參數.Length > 0)
                {
                    for (int 計算變數 = 0; 計算變數 < 參數.Length; 計算變數++)
                    {
                        資料庫命令.Parameters.Add(參數[計算變數]);
                    }
                }
                資料集合 = 資料庫命令.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception 例外訊息)
            {
                //寫入出錯Log檔
                資料庫命令.Cancel();
                資料集合.Close();
                關閉連線();
                throw new Exception(例外訊息.Message);
            }
            return 資料集合;
        }
        public SqlDataReader 取得SQL集合(string 指令)
        {
            SqlDataReader 資料集合 = null;
            try
            {
                打開連線();
                資料庫命令 = new SqlCommand(指令, 物件連線);
                資料集合 = 資料庫命令.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception 例外訊息)
            {
                //寫入出錯Log檔
                資料庫命令.Cancel();
                資料集合.Close();
                關閉連線();
                throw new Exception(例外訊息.Message);
            }
            return 資料集合;
        } 
    }
    public class 示範類別
    {

        public 示範類別()
        {
        }

        public string 讀取相關資料()
        {
            連線類別 連線類別物件 = new 連線類別();
            SqlDataReader 資料讀取器=連線類別物件.取得SQL集合("select 工號,姓名,地址,* from 人員資料表 where 工號='TA140015'");
            string 員工姓名 = "";
            try
            {
                if (資料讀取器.Read())
                {
                    if (!資料讀取器.IsDBNull(資料讀取器.GetOrdinal("姓名")))
                    {
                        員工姓名 = 資料讀取器.GetString(1);
                    }
                }
            }
            catch (Exception 例外)
            {
                throw 例外;
            }
            finally
            {
                資料讀取器.Close();
            }

            return 員工姓名;

        }
    }
}
