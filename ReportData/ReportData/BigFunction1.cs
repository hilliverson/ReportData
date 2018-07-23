using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportData
{
    class BigFunction1
    {
        public string GetAllEntity1(Hashtable hTableQuery)
        {
            try
            {
                DataSet objDS = new DataSet();
                StringBuilder strBuilder = new StringBuilder();
                //起始時間
                string strStartDate = string.Empty;
                //結束時間
                string strEndDate = string.Empty;
                //false ByWeek/true ByDate
                bool bolByDateWeek = false;

                //年
                string strYear = string.Empty;
                //第幾週
                int intWeek = 0;
                //第幾週的前幾週
                int intPriWeek = 0;
                //群組
                string strGroup = string.Empty;
                //BU
                string strBu = string.Empty;
                //Tech
                string strTech = string.Empty;
                string strSeq = "";

                string strUserCol = "CLOSE_ID";

                if (hTableQuery.ContainsKey("USER_COL"))
                {
                    strUserCol = hTableQuery["USER_COL"].ToString();
                }
                if (hTableQuery.ContainsKey("SEQ"))
                {
                    strSeq = hTableQuery["SEQ"].ToString();
                }
                //查詢前幾筆資料
                //int intAmount = 0;

                if (hTableQuery.ContainsKey("DATE_WEEK"))
                {
                    bolByDateWeek = (bool)hTableQuery["DATE_WEEK"];
                }

                if (hTableQuery.ContainsKey("START_DATE"))
                {
                    strStartDate = hTableQuery["START_DATE"].ToString().Trim();
                }

                if (hTableQuery.ContainsKey("END_DATE"))
                {
                    strEndDate = hTableQuery["END_DATE"].ToString();

                }
                if (hTableQuery.ContainsKey("YEAR"))
                {
                    strYear = hTableQuery["YEAR"].ToString().Trim();

                }
                if (hTableQuery.ContainsKey("WEEK"))
                {
                    intWeek = Convert.ToInt32(hTableQuery["WEEK"].ToString().Trim());
                }
                if (hTableQuery.ContainsKey("PRI_WEEK"))
                {
                    intPriWeek = Convert.ToInt32(hTableQuery["PRI_WEEK"].ToString().Trim());
                }
                if (intPriWeek > intWeek)
                {
                    intPriWeek = intWeek;
                }

                if (hTableQuery.ContainsKey("BU"))
                {
                    strBu = hTableQuery["BU"].ToString().Trim();
                }
                if (hTableQuery.ContainsKey("TECH"))
                {
                    strTech = hTableQuery["TECH"].ToString().Trim();
                }
                string strUserGroup = string.Empty;
                if (hTableQuery.ContainsKey("USER_GROUP"))
                {
                    strUserGroup = hTableQuery["USER_GROUP"].ToString();
                }

                SqlParameter[] paras = new SqlParameter[6];
                paras[0] = new SqlParameter("@START_DATE", SqlDbType.DateTime);
                if (!string.IsNullOrEmpty(strStartDate))
                {
                    paras[0].Value = Convert.ToDateTime(strStartDate + " 00:00:00");
                }
                else
                    paras[0].Value = DBNull.Value;

                paras[1] = new SqlParameter("@END_DATE", SqlDbType.DateTime);
                if (!string.IsNullOrEmpty(strEndDate))
                {
                    paras[1].Value = Convert.ToDateTime(strEndDate + " 23:59:59");
                }
                else
                    paras[1].Value = DBNull.Value;

                paras[2] = new SqlParameter("@GROUP", SqlDbType.NVarChar);
                paras[2].Value = strGroup;
                paras[3] = new SqlParameter("@BU_NAME", SqlDbType.NVarChar);
                paras[3].Value = strBu;
                paras[4] = new SqlParameter("@TECH", SqlDbType.NVarChar);
                paras[4].Value = strTech;
                paras[5] = new SqlParameter("@SEQ", SqlDbType.Int);
                paras[5].Value = strSeq;

                #region By Date
                if (bolByDateWeek)
                {
                    //查詢從BD_BU_PRODUCTLINE 查找BU_NAME構建表頭
                    string strTopSql = "select t.TYPE ";
                    string strSumCol = "";
                    string strSqlCmd = "select distinct a.BU_NAME,b.SEQ from BD_BU_PRODUCTLINE a " +
                        " left join BD_BU_LIST b on a.BU_NAME =b.BU_NAME where 1=1 ";
                    string strGroupSql = Public_Helper.GetGroupSqlByReport(strUserGroup);

                    //選擇的產品線不爲空 
                    if (!string.IsNullOrEmpty(strBu))
                    {
                        strSqlCmd += " and a.BU_NAME=@BU_NAME ";
                    }
                    strSqlCmd += " order by b.SEQ ";
                    objDS = c_SqlHelper.getDataSet(strSqlCmd, paras);
                    if (c_SqlHelper.chkDataSet(objDS))
                    {
                        int intColumn = objDS.Tables[0].Rows.Count + 2;

                        //構建表頭
                        strBuilder.Append(" <table  class='TableStyle' cellspacing='1' cellpadding='0' width='100%' border='0'>");
                        strBuilder.Append(" <tr bgcolor='#FF99CC' height='20px' ><td  align='right' >Period</td><td " +
                            " colspan='" + intColumn + "' align='left' >" + strStartDate + " - " + strEndDate + "</td></tr>");

                        strBuilder.Append(" <tr  ><td nowrap class='TableBgblue'  width='120px'  >BU</td><td nowrap class='TableBgblue'  width='100px'   >&nbsp;</td>");
                        strBuilder.Append("<td nowrap class='TableBgblue' >Total</td>");
                        ArrayList arrCol = new ArrayList();
                        for (int i = 0; i < objDS.Tables[0].Rows.Count; i++)
                        {
                            string strBuName = objDS.Tables[0].Rows[i]["BU_NAME"].ToString().Trim();
                            arrCol.Add(strBuName);
                            strTopSql += ",isnull([" + strBuName + "],0) as [" + strBuName + "]";
                            strSumCol += "," + "[" + strBuName + "]";

                            strBuilder.Append(" <td nowrap class='TableBgblue' >" + strBuName + "</td>");
                        }

                        strSumCol = strSumCol.Substring(1);

                        //表頭結束
                        strBuilder.Append("</tr>");

                        //部門的總通話記錄by date & by BU 2009/5/11 不包括異常資料
                        string strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                            " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                            " a.ISSUE_ID=b.ISSUE_ID  ";
                        if (!string.IsNullOrEmpty(strUserGroup))
                        {
                            strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                        }

                        strTempSql += "where  convert(nvarchar,a.ISSUE_DATE,111) " +
                " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) and EXCEPTION='N'  ";
                        //Seq不爲空則爲查找大單的資料
                        if (!string.IsNullOrEmpty(strSeq))
                        {
                            strTempSql += " and b.SEQ=@SEQ ";
                        }
                        //產品線
                        string strSqlCmd1 = "select PRODUCT_LINE from BD_BU_PRODUCTLINE where BU_NAME=@BU_NAME  ";

                        //Tech
                        if (!string.IsNullOrEmpty(strTech))
                        {
                            strTempSql += " and b.TECH=@TECH";
                        }

                        if (!string.IsNullOrEmpty(strBu))
                        {
                            strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                            //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                        }

                        //InBound資料 2009/5/11 不包括異常資料
                        string strInBoundSql = "select bb.BU_NAME,count(*) as CNT,'inbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                            " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME union all select '' BU_NAME,0 CNT,'inbound' as TYPE  ";

                        strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                           " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                           " a.ISSUE_ID=b.ISSUE_ID inner join IS_OUTBOUND_ITEM d on d.ISSUE_SN=b.ISSUE_SN  ";
                        if (!string.IsNullOrEmpty(strUserGroup))
                        {
                            strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                        }

                        strTempSql += "  where   convert(nvarchar,b.ACTION_START_DATE,111) between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)   ";
                        //Seq不爲空則爲查找大單的資料
                        if (!string.IsNullOrEmpty(strSeq))
                        {
                            strTempSql += " and b.SEQ=@SEQ ";
                        }

                        //Tech
                        if (!string.IsNullOrEmpty(strTech))
                        {
                            strTempSql += " and b.TECH=@TECH";
                        }

                        if (!string.IsNullOrEmpty(strBu))
                        {
                            strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                            //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                        }

                        //OutBound
                        string strOutBoundSql = "select bb.BU_NAME,count(*) as CNT,'outbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                             " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound' as TYPE ";

                        //回電記錄管理維護作業中的回電記錄
                        //2009/7/23 回電數量(回電記錄+單獨的回電記錄)
                        strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                           " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a inner join OB_ACTION_ITEM b on " +
                           " a.OUTBOUND_ID=b.OUTBOUND_ID ";
                        if (!string.IsNullOrEmpty(strUserGroup))
                        {
                            strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                        }

                        strTempSql += "where  convert(nvarchar,a.OUTBOUND_DATE,111) " +
                " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)  ";
                        //Seq不爲空則爲查找大單的資料
                        if (!string.IsNullOrEmpty(strSeq))
                        {
                            strTempSql += " and b.SEQ=@SEQ ";
                        }


                        //Tech
                        if (!string.IsNullOrEmpty(strTech))
                        {
                            strTempSql += " and b.TECH=@TECH";
                        }

                        if (!string.IsNullOrEmpty(strBu))
                        {
                            strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                            //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                        }

                        //單獨的回電記錄
                        strTempSql += " union all select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                           " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a " +
                           " inner join OB_ACTION_ITEM b on a.OUTBOUND_ID=b.OUTBOUND_ID  " +
                           " inner join OB_OUTBOUND_HEADER c on c.OUTBOUND_ID=a.OUTBOUND_ID " +
                           " inner join OB_OUTBOUND_ITEM d on c.OB_ID=d.OB_ID and d.OUTBOUND_SN=b.OUTBOUND_SN  ";

                        if (!string.IsNullOrEmpty(strUserGroup))
                        {
                            strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                        }

                        strTempSql += " where  convert(nvarchar,a.OUTBOUND_DATE,111) " +
                        " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)  ";
                        //Seq不爲空則爲查找大單的資料
                        if (!string.IsNullOrEmpty(strSeq))
                        {
                            strTempSql += " and b.SEQ=@SEQ ";
                        }


                        //Tech
                        if (!string.IsNullOrEmpty(strTech))
                        {
                            strTempSql += " and b.TECH=@TECH";
                        }

                        if (!string.IsNullOrEmpty(strBu))
                        {
                            strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                            //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                        }

                        //OutBound
                        string strOutBoundSql1 = "select bb.BU_NAME,count(*) as CNT,'outbound1' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                             " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound1' as TYPE ";

                        //統計異常資料2009/5/11
                        strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                         " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                         " a.ISSUE_ID=b.ISSUE_ID ";
                        if (!string.IsNullOrEmpty(strUserGroup))
                        {
                            strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                        }

                        strTempSql += "where  convert(nvarchar,a.ISSUE_DATE,111) " +
                         " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) and EXCEPTION='Y'  ";
                        //Seq不爲空則爲查找大單的資料
                        if (!string.IsNullOrEmpty(strSeq))
                        {
                            strTempSql += " and b.SEQ=@SEQ ";
                        }


                        //Tech
                        if (!string.IsNullOrEmpty(strTech))
                        {
                            strTempSql += " and b.TECH=@TECH";
                        }

                        if (!string.IsNullOrEmpty(strBu))
                        {
                            strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                            //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                        }

                        //異常資料
                        string strExceptionSql = "select bb.BU_NAME,count(*) as CNT,'p-exception' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                            " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME union all select '' BU_NAME,0 CNT,'p-exception' as TYPE  ";
                        //Summary
                        string strSummarySql = "select BU_NAME,sum(CNT),'summary' as TYPE from(" + strInBoundSql + " union all "
                            + strOutBoundSql + " union all " + strOutBoundSql1 + " union all " + strExceptionSql + ") bb " +
                            " group by BU_NAME ";

                        strSqlCmd = strTopSql + " from (" + strInBoundSql + " union all " + strOutBoundSql + " union all "
                            + strOutBoundSql1 + "union all " + strExceptionSql
                            + " union all " + strSummarySql + " ) dd pivot(sum(dd.CNT) " +
                            " for dd.BU_NAME in(" + strSumCol + ")) t order by 1 ";

                        DataSet objBU = c_SqlHelper.getDataSet(strSqlCmd, paras);
                        if (c_SqlHelper.chkDataSet(objBU))
                        {

                            for (int k = 0; k < objBU.Tables[0].Rows.Count; k++)
                            {
                                int intSumTotal = 0;
                                string strType = "";
                                switch (k.ToString())
                                {
                                    case "0":
                                        strType = "IN";
                                        strBuilder.Append("<tr bgcolor='#FAFA9A'><td rowspan='5' align='center' >Dept</td><td  align='center' >In</td>");
                                        break;
                                    case "1":
                                        strType = "OUT(1)";
                                        strBuilder.Append("<tr bgcolor='#FAFA9A' ><td align='center' >Out(1)</td>");
                                        break;
                                    case "2":
                                        strType = "OUT(2)";
                                        strBuilder.Append("<tr bgcolor='#FAFA9A' ><td align='center' >Out(2)</td>");
                                        break;
                                    case "3":
                                        strType = "EXCEPTION";
                                        strBuilder.Append("<tr bgcolor='#FAFA9A' ><td align='center' >Exception</td>");
                                        break;
                                    case "4":
                                        strType = "TOTAL";
                                        strBuilder.Append("<tr bgcolor='#FAFA9A'><td  align='center' >Total</td>");
                                        break;
                                }

                                //查找BU的資料Close id爲空
                                hTableQuery["CLOSE_ID"] = "";
                                //獲得來電總數
                                intSumTotal = GetBoundTotal(hTableQuery, strType);
                                if (intSumTotal == 0)
                                {
                                    strBuilder.Append("<td  style='height:22px;text-indent:2px;font-size:12px;' >" + intSumTotal + "</td>");
                                }
                                else
                                {
                                    strBuilder.Append("<td  style='color:Blue; text-decoration:underline;cursor:hand;height:22px;text-indent:2px;font-size:12px;' onclick=\"openTechQWindow('','"
                                          + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                          Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                          Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + +intSumTotal + "</td>");
                                }
                                for (int j = 0; j < arrCol.Count; j++)
                                {
                                    //intSumTotal += Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString());

                                    if (Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString()) > 0)
                                    {

                                        strBuilder.Append("<td  style='color:Blue; text-decoration:underline;cursor:hand;height:22px;text-indent:2px;font-size:12px;' onclick=\"openTechQWindow('" + Check_Helper.ConvertEnUrl(arrCol[j].ToString()) + "','"
                                          + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                          Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                          Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                    }
                                    else
                                    {
                                        strBuilder.Append("<td style='height:22px;text-indent:2px;font-size:12px;' >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                    }
                                }
                                strBuilder.Append("</tr>");

                            }
                        }


                        //ByDate By user 找出有哪些員工有通話紀錄
                        strSqlCmd = "select " + strUserCol + ", a.USER_NAME_ENG from AS_USER a inner join ( ";
                        strSqlCmd += " select distinct " + strUserCol + " from IS_ACTION_HEADER where 1=1 ";
                        //GroupSql
                        if (!string.IsNullOrEmpty(strGroupSql))
                        {
                            strSqlCmd += " and AGENT_ID in (" + strGroupSql + ") ";
                        }
                        strSqlCmd += " union select distinct " + strUserCol + " from OB_ACTION_HEADER where 1=1  ";
                        //GroupSql
                        if (!string.IsNullOrEmpty(strGroupSql))
                        {
                            strSqlCmd += " and AGENT_ID in (" + strGroupSql + ") ";
                        }
                        strSqlCmd += " ) as closer on a.user_id = closer." + strUserCol + " order by USER_NAME_ENG ";
                        DataSet objUser = c_SqlHelper.getDataSet(strSqlCmd, paras);
                        if (c_SqlHelper.chkDataSet(objUser))
                        {
                            for (int n = 0; n < objUser.Tables[0].Rows.Count; n++)
                            {
                                string strCloseID = objUser.Tables[0].Rows[n][strUserCol].ToString().Trim();

                                strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                         " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                         " a.ISSUE_ID=b.ISSUE_ID  ";
                                if (!string.IsNullOrEmpty(strUserGroup))
                                {
                                    strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                                }

                                strTempSql += " where  convert(nvarchar,a.ISSUE_DATE,111) " +
                               " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) and EXCEPTION='N'  " +
                                " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "' ";
                                //Seq不爲空則爲查找大單的資料
                                if (!string.IsNullOrEmpty(strSeq))
                                {
                                    strTempSql += " and b.SEQ=@SEQ ";
                                }

                                //Tech
                                if (!string.IsNullOrEmpty(strTech))
                                {
                                    strTempSql += " and b.TECH=@TECH";
                                }

                                if (!string.IsNullOrEmpty(strBu))
                                {
                                    strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                    //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                }

                                //InBound資料
                                strInBoundSql = "select bb.BU_NAME,count(*) as CNT,'inbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                   " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'inbound' as TYPE ";

                                strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                                   " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                                   " a.ISSUE_ID=b.ISSUE_ID inner join IS_OUTBOUND_ITEM d on d.ISSUE_SN=b.ISSUE_SN  ";
                                if (!string.IsNullOrEmpty(strUserGroup))
                                {
                                    strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                                }

                                strTempSql += "  where  convert(nvarchar,b.ACTION_START_DATE,111) between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)   " +
                                   " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "'";
                                //Seq不爲空則爲查找大單的資料
                                if (!string.IsNullOrEmpty(strSeq))
                                {
                                    strTempSql += " and b.SEQ=@SEQ ";
                                }

                                //Tech
                                if (!string.IsNullOrEmpty(strTech))
                                {
                                    strTempSql += " and b.TECH=@TECH";
                                }

                                if (!string.IsNullOrEmpty(strBu))
                                {
                                    strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                    //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                }

                                //OutBound
                                strOutBoundSql = "select bb.BU_NAME,count(*) as CNT,'outbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                    " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound' as TYPE  ";

                                //回電記錄資料
                                strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                       " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a inner join OB_ACTION_ITEM b on " +
                       " a.OUTBOUND_ID=b.OUTBOUND_ID ";

                                if (!string.IsNullOrEmpty(strUserGroup))
                                {
                                    strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                                }

                                strTempSql += "where convert(nvarchar,a.OUTBOUND_DATE,111) " +
                        " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)   " +
                       " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "' ";
                                //Seq不爲空則爲查找大單的資料
                                if (!string.IsNullOrEmpty(strSeq))
                                {
                                    strTempSql += " and b.SEQ=@SEQ ";
                                }

                                //Tech
                                if (!string.IsNullOrEmpty(strTech))
                                {
                                    strTempSql += " and b.TECH=@TECH";
                                }

                                if (!string.IsNullOrEmpty(strBu))
                                {
                                    strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                    //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                }


                                //單獨的回電記錄
                                strTempSql += " union all select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                                   " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a " +
                                   " inner join OB_ACTION_ITEM b on a.OUTBOUND_ID=b.OUTBOUND_ID  " +
                                   " inner join OB_OUTBOUND_HEADER c on c.OUTBOUND_ID=a.OUTBOUND_ID " +
                                   " inner join OB_OUTBOUND_ITEM d on c.OB_ID=d.OB_ID and d.OUTBOUND_SN=b.OUTBOUND_SN  ";
                                if (!string.IsNullOrEmpty(strUserGroup))
                                {
                                    strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                                }

                                strTempSql += " where  convert(nvarchar,a.OUTBOUND_DATE,111) " +
                                " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) " +
                                "  and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "'";
                                //Seq不爲空則爲查找大單的資料
                                if (!string.IsNullOrEmpty(strSeq))
                                {
                                    strTempSql += " and b.SEQ=@SEQ ";
                                }


                                //Tech
                                if (!string.IsNullOrEmpty(strTech))
                                {
                                    strTempSql += " and b.TECH=@TECH";
                                }

                                if (!string.IsNullOrEmpty(strBu))
                                {
                                    strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                    //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                }

                                strOutBoundSql1 = "select bb.BU_NAME,count(*) as CNT,'outbound1' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                 " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound1' as TYPE ";

                                #region 統計異常資料2009/5/11 By Closer
                                strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                                 " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                                 " a.ISSUE_ID=b.ISSUE_ID  ";

                                if (!string.IsNullOrEmpty(strUserGroup))
                                {
                                    strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "'";
                                }

                                strTempSql += "where  convert(nvarchar,a.ISSUE_DATE,111) " +
                        " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) and EXCEPTION='Y' " +
                                   " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "'";
                                //Seq不爲空則爲查找大單的資料
                                if (!string.IsNullOrEmpty(strSeq))
                                {
                                    strTempSql += " and b.SEQ=@SEQ ";
                                }


                                //Tech
                                if (!string.IsNullOrEmpty(strTech))
                                {
                                    strTempSql += " and b.TECH=@TECH";
                                }

                                if (!string.IsNullOrEmpty(strBu))
                                {
                                    strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                    //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                }

                                //異常資料
                                strExceptionSql = "select bb.BU_NAME,count(*) as CNT,'p-exception' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                   " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME union all select '' BU_NAME,0 CNT,'p-exception' as TYPE  ";
                                #endregion

                                //Summary
                                strSummarySql = "select BU_NAME,sum(CNT),'summary' as TYPE from(" + strInBoundSql + " union all "
                                    + strOutBoundSql + " union all " + strOutBoundSql1 + " union all " + strExceptionSql + " ) bb " +
                                   " group by BU_NAME ";

                                strSqlCmd = strTopSql + " from (" + strInBoundSql + " union all " + strOutBoundSql
                                    + " union all " + strOutBoundSql1 + " union all " + strExceptionSql + " union all " + strSummarySql + " ) dd pivot(sum(dd.CNT) " +
                                   " for dd.BU_NAME in(" + strSumCol + ")) t order by 1 ";

                                objBU = c_SqlHelper.getDataSet(strSqlCmd, paras);
                                if (c_SqlHelper.chkDataSet(objBU))
                                {

                                    for (int k = 0; k < objBU.Tables[0].Rows.Count; k++)
                                    {
                                        int intSumTotal = 0;
                                        string strType = "";
                                        switch (k.ToString())
                                        {
                                            case "0":
                                                strType = "IN";
                                                strBuilder.Append("<tr><td rowspan='5' class='TableBgblue'>" + new IssueOperate().GetUserEngName(strCloseID) + "</td><td class='TableBgblue'>In</td>");
                                                break;
                                            case "1":
                                                strType = "OUT(1)";
                                                strBuilder.Append("<tr><td class='TableBgblue'>Out(1)</td>");
                                                break;
                                            case "2":
                                                strType = "OUT(2)";
                                                strBuilder.Append("<tr><td class='TableBgblue'>Out(2)</td>");
                                                break;
                                            case "3":
                                                strType = "EXCEPTION";
                                                strBuilder.Append("<tr><td class='TableBgblue'>Exception</td>");
                                                break;
                                            case "4":
                                                strType = "TOTAL";
                                                strBuilder.Append("<tr  bgcolor='#ffcc99'><td align='center' >Total</td>");
                                                break;
                                        }
                                        hTableQuery["CLOSE_ID"] = strCloseID;
                                        //獲得來電總數
                                        intSumTotal = GetBoundTotal(hTableQuery, strType);

                                        if (k == 4)
                                        {
                                            if (intSumTotal == 0)
                                            {
                                                strBuilder.Append("<td style='height:22px;text-indent:2px;font-size:12px;'>" + intSumTotal + "</td>");
                                            }
                                            else
                                            {
                                                strBuilder.Append("<td style='height:22px;text-indent:2px;font-size:12px;color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('','"
                                                   + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                   Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                   Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + intSumTotal + "</td>");
                                            }
                                        }
                                        else
                                        {
                                            if (intSumTotal == 0)
                                            {
                                                strBuilder.Append("<td class='TableBgGray'>" + intSumTotal + "</td>");
                                            }
                                            else
                                            {
                                                strBuilder.Append("<td class='TableBgGray' style='color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('','"
                                                   + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                   Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                   Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + intSumTotal + "</td>");
                                            }
                                        }

                                        for (int j = 0; j < arrCol.Count; j++)
                                        {

                                            //Total顯示行加顔色
                                            if (k == 4)
                                            {
                                                if (Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString()) > 0)
                                                {

                                                    strBuilder.Append("<td style='height:22px;text-indent:2px;font-size:12px;color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('" + Check_Helper.ConvertEnUrl(arrCol[j].ToString()) + "','"
                                                      + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                      Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                      Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                }
                                                else
                                                {
                                                    strBuilder.Append("<td  style='height:22px;text-indent:2px;font-size:12px;'>" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                }
                                            }
                                            else
                                            {
                                                if (Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString()) > 0)
                                                {

                                                    strBuilder.Append("<td class='TableBgGray' style='color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('" + Check_Helper.ConvertEnUrl(arrCol[j].ToString()) + "','"
                                                      + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                      Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                      Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                }
                                                else
                                                {
                                                    strBuilder.Append("<td  class='TableBgGray'>" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                }
                                            }

                                        }
                                        strBuilder.Append("</tr>");

                                    }
                                }

                            }
                        }
                        strBuilder.Append("</table>");
                    }
                    else
                    {
                        return "";
                    }
                }
                #endregion

                #region By Week
                if (!bolByDateWeek)
                {
                    //依查詢條件，若為2009年第六週的前2週，先取出第六週前兩週的起始及結束日
                    int intStartWeek = intWeek - intPriWeek + 1;
                    string strSqlCmd = "select a.Recordweek,convert(nvarchar ,Min(a.Recorddate),111) as  startdate," +
                         "convert(nvarchar ,Max(a.RecordDate),111) as enddate,RecordWeek " +
                        " from bd_calendar  a left join bd_week  b on a.week_name = b.week_name " +
                        " where a.recordyear = b.year and a.week_name =N'" + strYear + "' and " +
                        " a.recordweek between " + intStartWeek + " and   " + intWeek + " group by Recordweek ";
                    string strGroupSql = Public_Helper.GetGroupSqlByReport(strUserGroup);

                    DataSet objDSWeek = c_SqlHelper.getDataSet(strSqlCmd);
                    for (int m = 0; m < objDSWeek.Tables[0].Rows.Count; m++)
                    {
                        strStartDate = objDSWeek.Tables[0].Rows[m]["startdate"].ToString().Trim();
                        strEndDate = objDSWeek.Tables[0].Rows[m]["enddate"].ToString().Trim();
                        paras[0].Value = Convert.ToDateTime(strStartDate + " 00:00:00");
                        paras[1].Value = Convert.ToDateTime(strEndDate + " 23:59:59");
                        string strRecordWeek = objDSWeek.Tables[0].Rows[m]["RecordWeek"].ToString().Trim();
                        //查詢從BD_BU_PRODUCTLINE 查找BU_NAME構建表頭
                        string strTopSql = "select t.TYPE ";
                        string strSumCol = "";
                        strSqlCmd = "select distinct a.BU_NAME,b.SEQ from BD_BU_PRODUCTLINE a " +
                       " left join BD_BU_LIST b on a.BU_NAME =b.BU_NAME where 1=1 ";

                        //選擇的產品線不爲空 
                        if (!string.IsNullOrEmpty(strBu))
                        {
                            strSqlCmd += " and a.BU_NAME=@BU_NAME ";
                        }
                        strSqlCmd += " order by b.SEQ ";
                        objDS = c_SqlHelper.getDataSet(strSqlCmd, paras);
                        if (c_SqlHelper.chkDataSet(objDS))
                        {
                            int intColumn = objDS.Tables[0].Rows.Count + 2;

                            //構建表頭
                            strBuilder.Append(" <table  class='TableStyle' cellspacing='1' cellpadding='0' width='100%' border='0'>");
                            strBuilder.Append(" <tr bgcolor='#FF99CC' height='20px' ><td  align='right' >Period</td><td " +
                                " colspan='" + intColumn + "' align='left' >" + strStartDate + " - " + strEndDate + "(" + strRecordWeek + ")</td></tr>");

                            strBuilder.Append(" <tr><td nowrap class='TableBgblue' width='120px' >BU</td><td nowrap class='TableBgblue' width='100px' >&nbsp;</td>");
                            strBuilder.Append("<td nowrap class='TableBgblue' >Total</td>");
                            ArrayList arrCol = new ArrayList();
                            for (int i = 0; i < objDS.Tables[0].Rows.Count; i++)
                            {
                                string strBuName = objDS.Tables[0].Rows[i]["BU_NAME"].ToString().Trim();
                                arrCol.Add(strBuName);
                                strTopSql += ",isnull([" + strBuName + "],0) as [" + strBuName + "]";
                                strSumCol += "," + "[" + strBuName + "]";

                                strBuilder.Append(" <td nowrap class='TableBgblue' >" + strBuName + "</td>");
                            }
                            strSumCol = strSumCol.Substring(1);

                            //表頭結束
                            strBuilder.Append("</tr>");

                            //部門的總通話記錄by date & by BU
                            string strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                                " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                                " a.ISSUE_ID=b.ISSUE_ID  ";
                            if (!string.IsNullOrEmpty(strUserGroup))
                            {
                                strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                            }

                            strTempSql += "where  convert(nvarchar,a.ISSUE_DATE,111) " +
                         " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) and EXCEPTION='N' ";
                            //Seq不爲空則爲查找大單的資料
                            if (!string.IsNullOrEmpty(strSeq))
                            {
                                strTempSql += " and b.SEQ=@SEQ ";
                            }
                            //產品線
                            string strSqlCmd1 = "select PRODUCT_LINE from BD_BU_PRODUCTLINE where BU_NAME=@BU_NAME  ";

                            //Tech
                            if (!string.IsNullOrEmpty(strTech))
                            {
                                strTempSql += " and b.TECH=@TECH";
                            }

                            if (!string.IsNullOrEmpty(strBu))
                            {
                                strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                            }

                            //InBound資料
                            string strInBoundSql = "select bb.BU_NAME,count(*) as CNT,'inbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME union all select '' BU_NAME,0 CNT,'inbound' as TYPE  ";

                            strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                               " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                               " a.ISSUE_ID=b.ISSUE_ID inner join IS_OUTBOUND_ITEM d on d.ISSUE_SN=b.ISSUE_SN  ";

                            if (!string.IsNullOrEmpty(strUserGroup))
                            {
                                strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                            }

                            strTempSql += "  where  b.ACTION_START_DATE between @START_DATE and @END_DATE   ";
                            //Seq不爲空則爲查找大單的資料
                            if (!string.IsNullOrEmpty(strSeq))
                            {
                                strTempSql += " and b.SEQ=@SEQ ";
                            }


                            //Tech
                            if (!string.IsNullOrEmpty(strTech))
                            {
                                strTempSql += " and b.TECH=@TECH";
                            }

                            if (!string.IsNullOrEmpty(strBu))
                            {
                                strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                            }


                            //OutBound
                            string strOutBoundSql = "select bb.BU_NAME,count(*) as CNT,'outbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                 " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound' as TYPE ";

                            //回電記錄管理維護作業中的回電記錄
                            strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                               " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a inner join OB_ACTION_ITEM b on " +
                               " a.OUTBOUND_ID=b.OUTBOUND_ID ";
                            if (!string.IsNullOrEmpty(strUserGroup))
                            {
                                strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                            }

                            strTempSql += " where  convert(nvarchar,a.OUTBOUND_DATE,111) " +
                         " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)  ";
                            //Seq不爲空則爲查找大單的資料
                            if (!string.IsNullOrEmpty(strSeq))
                            {
                                strTempSql += " and b.SEQ=@SEQ ";
                            }


                            //Tech
                            if (!string.IsNullOrEmpty(strTech))
                            {
                                strTempSql += " and b.TECH=@TECH";
                            }

                            if (!string.IsNullOrEmpty(strBu))
                            {
                                strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                            }

                            //單獨的回電記錄
                            strTempSql += " union all select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                               " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a " +
                               " inner join OB_ACTION_ITEM b on a.OUTBOUND_ID=b.OUTBOUND_ID  " +
                               " inner join OB_OUTBOUND_HEADER c on c.OUTBOUND_ID=a.OUTBOUND_ID " +
                               " inner join OB_OUTBOUND_ITEM d on c.OB_ID=d.OB_ID and d.OUTBOUND_SN=b.OUTBOUND_SN  ";

                            if (!string.IsNullOrEmpty(strUserGroup))
                            {
                                strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                            }

                            strTempSql += " where  convert(nvarchar,a.OUTBOUND_DATE,111) " +
                            " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) ";

                            //Seq不爲空則爲查找大單的資料
                            if (!string.IsNullOrEmpty(strSeq))
                            {
                                strTempSql += " and b.SEQ=@SEQ ";
                            }


                            //Tech
                            if (!string.IsNullOrEmpty(strTech))
                            {
                                strTempSql += " and b.TECH=@TECH";
                            }

                            if (!string.IsNullOrEmpty(strBu))
                            {
                                strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";

                            }

                            //OutBound
                            string strOutBoundSql1 = "select bb.BU_NAME,count(*) as CNT,'outbound1' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                 " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound1' as TYPE ";
                            //統計異常資料2009/5/11
                            strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                             " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                             " a.ISSUE_ID=b.ISSUE_ID ";
                            if (!string.IsNullOrEmpty(strUserGroup))
                            {
                                strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                            }

                            strTempSql += "  where  convert(nvarchar,a.ISSUE_DATE,111) " +
                         " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) and EXCEPTION='Y'  ";
                            //Seq不爲空則爲查找大單的資料
                            if (!string.IsNullOrEmpty(strSeq))
                            {
                                strTempSql += " and b.SEQ=@SEQ ";
                            }

                            ////群組 
                            //if (!string.IsNullOrEmpty(strGroup))
                            //{
                            //    strTempSql += " and c.COUNTRY=@GROUP";
                            //}
                            //Tech
                            if (!string.IsNullOrEmpty(strTech))
                            {
                                strTempSql += " and b.TECH=@TECH";
                            }

                            if (!string.IsNullOrEmpty(strBu))
                            {
                                strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                            }

                            //異常資料
                            string strExceptionSql = "select bb.BU_NAME,count(*) as CNT,'p-exception' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME union all select '' BU_NAME,0 CNT,'p-exception' as TYPE  ";
                            //Summary
                            string strSummarySql = "select BU_NAME,sum(CNT),'summary' as TYPE from(" + strInBoundSql + " union all "
                                + strOutBoundSql + " union all " + strOutBoundSql1 + " union all " + strExceptionSql + ") bb " +
                                " group by BU_NAME ";

                            strSqlCmd = strTopSql + " from (" + strInBoundSql + " union all " + strOutBoundSql + " union all "
                                + strOutBoundSql1 + "union all " + strExceptionSql
                                + " union all " + strSummarySql + " ) dd pivot(sum(dd.CNT) " +
                                " for dd.BU_NAME in(" + strSumCol + ")) t order by 1 ";

                            DataSet objBU = c_SqlHelper.getDataSet(strSqlCmd, paras);
                            if (c_SqlHelper.chkDataSet(objBU))
                            {

                                for (int k = 0; k < objBU.Tables[0].Rows.Count; k++)
                                {
                                    int intSumTotal = 0;
                                    string strType = "";
                                    switch (k.ToString())
                                    {
                                        case "0":
                                            strType = "IN";
                                            strBuilder.Append("<tr bgcolor='#FAFA9A'><td rowspan='5' align='center' >Dept</td><td  align='center' >In</td>");
                                            break;
                                        case "1":
                                            strType = "OUT(1)";
                                            strBuilder.Append("<tr bgcolor='#FAFA9A' ><td align='center' >Out(1)</td>");
                                            break;
                                        case "2":
                                            strType = "OUT(2)";
                                            strBuilder.Append("<tr bgcolor='#FAFA9A' ><td align='center' >Out(2)</td>");
                                            break;
                                        case "3":
                                            strType = "EXCEPTION";
                                            strBuilder.Append("<tr bgcolor='#FAFA9A' ><td align='center' >Exception</td>");
                                            break;
                                        case "4":
                                            strType = "TOTAL";
                                            strBuilder.Append("<tr bgcolor='#FAFA9A'><td  align='center' >Total</td>");
                                            break;
                                    }

                                    hTableQuery["START_DATE"] = strStartDate;
                                    hTableQuery["END_DATE"] = strEndDate;
                                    hTableQuery["CLOSE_ID"] = "";
                                    //獲得來電總數
                                    intSumTotal = GetBoundTotal(hTableQuery, strType);
                                    if (intSumTotal == 0)
                                    {
                                        strBuilder.Append("<td  style='height:22px;text-indent:2px;font-size:12px;' >" + intSumTotal + "</td>");
                                    }
                                    else
                                    {
                                        strBuilder.Append("<td  style='color:Blue; text-decoration:underline;cursor:hand;height:22px;text-indent:2px;font-size:12px;' onclick=\"openTechQWindow('','"
                                              + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                              Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                              Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + +intSumTotal + "</td>");
                                    }
                                    for (int j = 0; j < arrCol.Count; j++)
                                    {
                                        //intSumTotal += Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString());

                                        if (Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString()) > 0)
                                        {

                                            strBuilder.Append("<td  style='color:Blue; text-decoration:underline;cursor:hand;height:22px;text-indent:2px;font-size:12px;' onclick=\"openTechQWindow('" + Check_Helper.ConvertEnUrl(arrCol[j].ToString()) + "','"
                                              + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                              Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                              Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                        }
                                        else
                                        {
                                            strBuilder.Append("<td style='height:22px;text-indent:2px;font-size:12px;' >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                        }
                                    }
                                    strBuilder.Append("</tr>");

                                }
                            }

                            //ByDate By user 找出有哪些員工有通話紀錄
                            strSqlCmd = "select " + strUserCol + ", a.USER_NAME_ENG from AS_USER a inner join ( ";
                            strSqlCmd += " select distinct " + strUserCol + " from IS_ACTION_HEADER where 1=1 ";
                            //GroupSql
                            if (!string.IsNullOrEmpty(strGroupSql))
                            {
                                strSqlCmd += " and AGENT_ID in (" + strGroupSql + ") ";
                            }
                            strSqlCmd += " union select distinct " + strUserCol + " from OB_ACTION_HEADER where 1=1  ";
                            //GroupSql
                            if (!string.IsNullOrEmpty(strGroupSql))
                            {
                                strSqlCmd += " and AGENT_ID in (" + strGroupSql + ") ";
                            }
                            strSqlCmd += " ) as closer on a.user_id = closer." + strUserCol + " order by USER_NAME_ENG ";

                            DataSet objUser = c_SqlHelper.getDataSet(strSqlCmd, paras);
                            if (c_SqlHelper.chkDataSet(objUser))
                            {
                                for (int n = 0; n < objUser.Tables[0].Rows.Count; n++)
                                {
                                    string strCloseID = objUser.Tables[0].Rows[n][strUserCol].ToString().Trim();

                                    strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                             " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                             " a.ISSUE_ID=b.ISSUE_ID ";
                                    if (!string.IsNullOrEmpty(strUserGroup))
                                    {
                                        strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                                    }

                                    strTempSql += " where  convert(nvarchar,a.ISSUE_DATE,111) " +
                         " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)  and EXCEPTION='N'  " +
                              " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "' ";
                                    //Seq不爲空則爲查找大單的資料
                                    if (!string.IsNullOrEmpty(strSeq))
                                    {
                                        strTempSql += " and b.SEQ=@SEQ ";
                                    }
                                    ////群組 
                                    //if (!string.IsNullOrEmpty(strGroup))
                                    //{
                                    //    strTempSql += " and c.COUNTRY=@GROUP";
                                    //}
                                    //Tech
                                    if (!string.IsNullOrEmpty(strTech))
                                    {
                                        strTempSql += " and b.TECH=@TECH";
                                    }

                                    if (!string.IsNullOrEmpty(strBu))
                                    {
                                        strTempSql += " and right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                        //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                    }

                                    //InBound資料
                                    strInBoundSql = "select bb.BU_NAME,count(*) as CNT,'inbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                       " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'inbound' as TYPE ";

                                    strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                                       " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                                       " a.ISSUE_ID=b.ISSUE_ID inner join IS_OUTBOUND_ITEM d on d.ISSUE_SN=b.ISSUE_SN  ";

                                    if (!string.IsNullOrEmpty(strUserGroup))
                                    {
                                        strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                                    }

                                    strTempSql += "  where  b.ACTION_START_DATE between @START_DATE and @END_DATE " +
                                       " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "'";
                                    //Seq不爲空則爲查找大單的資料
                                    if (!string.IsNullOrEmpty(strSeq))
                                    {
                                        strTempSql += " and b.SEQ=@SEQ ";
                                    }
                                    ////群組 
                                    //if (!string.IsNullOrEmpty(strGroup))
                                    //{
                                    //    strTempSql += " and c.COUNTRY=@GROUP";
                                    //}
                                    //Tech
                                    if (!string.IsNullOrEmpty(strTech))
                                    {
                                        strTempSql += " and b.TECH=@TECH";
                                    }

                                    if (!string.IsNullOrEmpty(strBu))
                                    {
                                        strTempSql += " and right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                        //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                    }

                                    //OutBound
                                    strOutBoundSql = "select bb.BU_NAME,count(*) as CNT,'outbound' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                        " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound' as TYPE  ";

                                    //回電記錄資料
                                    strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                           " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a inner join OB_ACTION_ITEM b on " +
                           " a.OUTBOUND_ID=b.OUTBOUND_ID ";
                                    if (!string.IsNullOrEmpty(strUserGroup))
                                    {
                                        strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                                    }

                                    strTempSql += " where  convert(nvarchar,a.OUTBOUND_DATE,111) " +
                         " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111)   " +
                            " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "' ";
                                    //Seq不爲空則爲查找大單的資料
                                    if (!string.IsNullOrEmpty(strSeq))
                                    {
                                        strTempSql += " and b.SEQ=@SEQ ";
                                    }

                                    //Tech
                                    if (!string.IsNullOrEmpty(strTech))
                                    {
                                        strTempSql += " and b.TECH=@TECH";
                                    }

                                    if (!string.IsNullOrEmpty(strBu))
                                    {
                                        strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                        //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                    }

                                    //單獨的回電記錄
                                    strTempSql += " union all select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                                       " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from OB_ACTION_HEADER a " +
                                       " inner join OB_ACTION_ITEM b on a.OUTBOUND_ID=b.OUTBOUND_ID  " +
                                       " inner join OB_OUTBOUND_HEADER c on c.OUTBOUND_ID=a.OUTBOUND_ID " +
                                       " inner join OB_OUTBOUND_ITEM d on c.OB_ID=d.OB_ID and d.OUTBOUND_SN=b.OUTBOUND_SN  ";
                                    if (!string.IsNullOrEmpty(strUserGroup))
                                    {
                                        strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                                    }

                                    strTempSql += " where  convert(nvarchar,a.OUTBOUND_DATE,111) " +
                                    " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) " +
                                    "  and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "'";
                                    //Seq不爲空則爲查找大單的資料
                                    if (!string.IsNullOrEmpty(strSeq))
                                    {
                                        strTempSql += " and b.SEQ=@SEQ ";
                                    }


                                    //Tech
                                    if (!string.IsNullOrEmpty(strTech))
                                    {
                                        strTempSql += " and b.TECH=@TECH";
                                    }

                                    if (!string.IsNullOrEmpty(strBu))
                                    {
                                        strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                        //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                    }

                                    strOutBoundSql1 = "select bb.BU_NAME,count(*) as CNT,'outbound1' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                     " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME  union all select '' BU_NAME,0 CNT,'outbound1' as TYPE ";

                                    #region 統計異常資料2009/5/11 By Closer
                                    strTempSql = "select case when len(b.PRODUCT_LINE)>2 then substring(b.PRODUCT_LINE ,charindex(',',b.PRODUCT_LINE)+1," +
                                     " len(b.PRODUCT_LINE)) else b.PRODUCT_LINE end as PRODUCT_LINE from IS_ACTION_HEADER a inner join IS_ACTION_ITEM b on " +
                                     " a.ISSUE_ID=b.ISSUE_ID  ";
                                    if (!string.IsNullOrEmpty(strUserGroup))
                                    {
                                        strTempSql += " inner join AS_USER e on e.USER_ID=a.AGENT_ID and e.GROUP_ID =N'" + strUserGroup + "' ";
                                    }

                                    strTempSql += " where  convert(nvarchar,a.ISSUE_DATE,111) " +
                         " between convert(nvarchar,@START_DATE,111) and convert(nvarchar,@END_DATE,111) and EXCEPTION='Y' " +
                                        " and a." + strUserCol + " =N'" + Check_Helper.ConvertFormat(strCloseID, 5) + "'";
                                    //Seq不爲空則爲查找大單的資料
                                    if (!string.IsNullOrEmpty(strSeq))
                                    {
                                        strTempSql += " and b.SEQ=@SEQ ";
                                    }


                                    //Tech
                                    if (!string.IsNullOrEmpty(strTech))
                                    {
                                        strTempSql += " and b.TECH=@TECH";
                                    }

                                    if (!string.IsNullOrEmpty(strBu))
                                    {
                                        strTempSql += " and  right(b.PRODUCT_LINE,len(b.PRODUCT_LINE)-charindex(',',b.PRODUCT_LINE)) in(" + strSqlCmd1 + ")";
                                        //" or PRODUCT_ID in (" + strSqlCmd1 + ")) ";
                                    }

                                    //異常資料
                                    strExceptionSql = "select bb.BU_NAME,count(*) as CNT,'p-exception' as TYPE from (" + strTempSql + ") aa left join BD_BU_PRODUCTLINE bb " +
                                       " on aa.PRODUCT_LINE=bb.PRODUCT_LINE group by bb.BU_NAME union all select '' BU_NAME,0 CNT,'p-exception' as TYPE  ";
                                    #endregion

                                    //Summary
                                    strSummarySql = "select BU_NAME,sum(CNT),'summary' as TYPE from(" + strInBoundSql + " union all "
                                        + strOutBoundSql + " union all " + strOutBoundSql1 + " union all " + strExceptionSql + " ) bb " +
                                       " group by BU_NAME ";

                                    strSqlCmd = strTopSql + " from (" + strInBoundSql + " union all " + strOutBoundSql
                                        + " union all " + strOutBoundSql1 + " union all " + strExceptionSql + " union all " + strSummarySql + " ) dd pivot(sum(dd.CNT) " +
                                       " for dd.BU_NAME in(" + strSumCol + ")) t order by 1 ";

                                    objBU = c_SqlHelper.getDataSet(strSqlCmd, paras);
                                    if (c_SqlHelper.chkDataSet(objBU))
                                    {

                                        for (int k = 0; k < objBU.Tables[0].Rows.Count; k++)
                                        {
                                            int intSumTotal = 0;
                                            string strType = "";
                                            switch (k.ToString())
                                            {
                                                case "0":
                                                    strType = "IN";
                                                    strBuilder.Append("<tr><td rowspan='5' class='TableBgblue'>" + new IssueOperate().GetUserEngName(strCloseID) + "</td><td class='TableBgblue'>In</td>");
                                                    break;
                                                case "1":
                                                    strType = "OUT(1)";
                                                    strBuilder.Append("<tr><td class='TableBgblue'>Out(1)</td>");
                                                    break;
                                                case "2":
                                                    strType = "OUT(2)";
                                                    strBuilder.Append("<tr><td class='TableBgblue'>Out(2)</td>");
                                                    break;
                                                case "3":
                                                    strType = "EXCEPTION";
                                                    strBuilder.Append("<tr><td class='TableBgblue'>Exception</td>");
                                                    break;
                                                case "4":
                                                    strType = "TOTAL";
                                                    strBuilder.Append("<tr  bgcolor='#ffcc99'><td align='center' >Total</td>");
                                                    break;
                                            }
                                            hTableQuery["START_DATE"] = strStartDate;
                                            hTableQuery["END_DATE"] = strEndDate;
                                            hTableQuery["CLOSE_ID"] = strCloseID;
                                            //獲得來電總數
                                            intSumTotal = GetBoundTotal(hTableQuery, strType);
                                            if (k == 4)
                                            {
                                                if (intSumTotal == 0)
                                                {
                                                    strBuilder.Append("<td style='height:22px;text-indent:2px;font-size:12px;'>" + intSumTotal + "</td>");
                                                }
                                                else
                                                {
                                                    strBuilder.Append("<td style='height:22px;text-indent:2px;font-size:12px;color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('','"
                                                       + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                       Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                       Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + intSumTotal + "</td>");
                                                }
                                            }
                                            else
                                            {
                                                if (intSumTotal == 0)
                                                {
                                                    strBuilder.Append("<td class='TableBgGray'>" + intSumTotal + "</td>");
                                                }
                                                else
                                                {
                                                    strBuilder.Append("<td class='TableBgGray' style='color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('','"
                                                       + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                       Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                       Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + intSumTotal + "</td>");
                                                }
                                            }
                                            for (int j = 0; j < arrCol.Count; j++)
                                            {
                                                //Total顯示行加顔色
                                                if (k == 4)
                                                {
                                                    if (Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString()) > 0)
                                                    {

                                                        strBuilder.Append("<td style='color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('" + Check_Helper.ConvertEnUrl(arrCol[j].ToString()) + "','"
                                                          + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                          Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                          Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                    }
                                                    else
                                                    {
                                                        strBuilder.Append("<td  >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                    }
                                                }
                                                else
                                                {
                                                    if (Convert.ToInt32(objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString()) > 0)
                                                    {

                                                        strBuilder.Append("<td class='TableBgGray' style='color:Blue; text-decoration:underline;cursor:hand;' onclick=\"openTechQByUser('" + Check_Helper.ConvertEnUrl(arrCol[j].ToString()) + "','"
                                                          + Check_Helper.ConvertEnUrl(strCloseID) + "','" + Check_Helper.ConvertEnUrl(strStartDate) + "','" +
                                                          Check_Helper.ConvertEnUrl(strEndDate) + "','" + Check_Helper.ConvertEnUrl(strGroup) + "','" +
                                                          Check_Helper.ConvertEnUrl(strBu) + "','" + Check_Helper.ConvertEnUrl(strTech) + "','" + Check_Helper.ConvertEnUrl(strType) + "','" + Check_Helper.ConvertEnUrl(strUserCol) + "','" + Check_Helper.ConvertEnUrl(strUserGroup) + "');\"  >" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                    }
                                                    else
                                                    {
                                                        strBuilder.Append("<td  class='TableBgGray'>" + objBU.Tables[0].Rows[k][arrCol[j].ToString()].ToString() + "</td>");
                                                    }
                                                }
                                            }
                                            strBuilder.Append("</tr>");
                                        }
                                    }

                                }
                            }
                            strBuilder.Append("</table>");
                        }
                    }
                }

                #endregion
                return strBuilder.ToString();
            }
            catch (Exception Exe)
            {
                throw new Exception(Exe.Message);
            }
        }

    }
}
