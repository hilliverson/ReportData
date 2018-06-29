using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportData
{
    class BigFunction2
    {
        private void SubmitAlarm(string strIndex, bool AddAlarm)
        {
            //逾時處理
            if (object.Equals(null, Session["g_UserID"]))
            {
                //定位到登錄頁面            
                Public_Helper.CallJavascriptEvent(this, "LoginOut", "GoToDefaultUrl();", true);
            }
            try
            {
                OutBoundEntity objOutBoundEntity = new OutBoundEntity();
                OutBoundOperate objOutBoundOpe = new OutBoundOperate();
                IssueOperate objIssueOperate = new IssueOperate();
                #region Alarm

                RadioButtonList radAlarmType = (RadioButtonList)Page.FindControl("radAlarmType" + strIndex);
                string strAlarmType = radAlarmType.SelectedValue.Trim();
                string strDate = ((TextBox)Page.FindControl("txtDateTime" + strIndex)).Text.Trim();


                HtmlInputHidden hidAlarmID = (HtmlInputHidden)Page.FindControl("hidAlarmID" + strIndex);
                string strAlarmID = hidAlarmID.Value.Trim();
                string strAttenTo = ((HtmlInputHidden)Page.FindControl("hidAttenTo" + strIndex)).Value.Trim();
                string strAttenCC = ((HtmlInputHidden)Page.FindControl("hidAttenCC" + strIndex)).Value.Trim();
                string strTO = ((TextBox)Page.FindControl("txtTo" + strIndex)).Text.Trim();
                string strCC = ((TextBox)Page.FindControl("txtCC" + strIndex)).Text.Trim();
                string strSubject = "";

                strSubject = ((TextBox)Page.FindControl("txtSubject" + strIndex)).Text.Trim();

                //同步Issue
                string strUserInfo = ((TextBox)Page.FindControl("txtUserInfo")).Text.Trim();
                string strUserTel = ((TextBox)Page.FindControl("txtTel1")).Text.Trim();

                #endregion

                #region Detail Record
                HtmlInputHidden hidOutBoundSn = (HtmlInputHidden)Page.FindControl("hidOutBoundSn" + strIndex);
                string strOutBoundSn = hidOutBoundSn.Value.Trim();
                string strChooseProductLine = "";
                RadioButtonList radProductLine = (RadioButtonList)Page.FindControl("radProductLine" + strIndex);

                string strProductLine = radProductLine.SelectedValue.ToString();
                if (string.IsNullOrEmpty(strProductLine))
                {
                    strProductLine = ((HtmlInputHidden)Page.FindControl("hidProductLine" + strIndex)).Value.Trim();
                    //20140604 hill add code for Mina來信要求要把產品線的預設值拿掉，所以程式必修增加判斷check 使用者是否有輸入產品線
                    Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["radProductLine"].ToString() + "');", true);
                    return;
                }
                strChooseProductLine = strProductLine;
                //有果有選中子產品線
                string strSubProcuct = ((HtmlInputHidden)Page.FindControl("hidSubProductLine" + strIndex)).Value.Trim();
                #region 產品線-子產品必填檢核
                //20120622_Claire 增加OutBound子產品線檢核
                //20110511 weihua.shen 卡關功能 加入產品線若為電腦週邊配備 則子產品為必填
                //20110929 卡關功能在增加 桌上型電腦、電腦零組件、數位家庭、工作站/伺服器、企業級資訊設備、Networking & Communication
                //20120622_Claire 卡關增加 桌上型電腦strProductLine == "6"
                //20121127_Claire 卡關增加 平板電腦strProductLine =="15"
                if (strProductLine == "6" || strProductLine == "7" || strProductLine == "8" || strProductLine == "10" || strProductLine == "11" ||
                    strProductLine == "12" || strProductLine == "13" || strProductLine == "15")
                {
                    //子產品不得為0
                    if (strSubProcuct == "")
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert",
                                                          "PrivateAlertMsg('" +
                                                          ViewState["valxSubProduct"].ToString() + "');", true);
                        return;
                    }
                }

                #endregion

                if (!string.IsNullOrEmpty(strSubProcuct))
                {
                    strProductLine += "," + strSubProcuct;
                    strChooseProductLine = strSubProcuct;
                }
                string strRelateID = ((TextBox)Page.FindControl("txtRelateID" + strIndex)).Text.Trim();
                string strModelID = ((TextBox)Page.FindControl("txtModel" + strIndex)).Text.Trim();
                string strClassValue = ((HtmlInputHidden)Page.FindControl("hidClassValue" + strIndex)).Value.Trim();
                string strProductID = ((HtmlInputHidden)Page.FindControl("hidProductID" + strIndex)).Value.Trim();
                string strProductVersion = "";
                //Model可修改時檢核
                if (((TextBox)Page.FindControl("txtModel" + strIndex)).Enabled)
                {
                    //檢核輸入的Model 是否已聯接地址一至
                    objIssueOperate.CheckModelNameLink(ref strModelID, ref strProductVersion, ref strClassValue, ref strProductID);
                }
                ((HtmlInputHidden)Page.FindControl("hidClassValue" + strIndex)).Value = strClassValue;
                ((HtmlInputHidden)Page.FindControl("hidProductID" + strIndex)).Value = strProductID;
                //if (strModelID.IndexOf("(") != -1)
                //{
                //    strProductVersion = strModelID.Substring(strModelID.IndexOf("(") + 1, strModelID.Length - strModelID.IndexOf("(") - 2);
                //}

                string strRmaNo = ((TextBox)Page.FindControl("txtRmaNo" + strIndex)).Text.Trim();

                RadioButtonList radAction = (RadioButtonList)Page.FindControl("radAction" + strIndex);
                string strAction = radAction.SelectedValue.Trim();

                RadioButtonList radTech = (RadioButtonList)Page.FindControl("radTech" + strIndex);
                string strTech = radTech.SelectedValue.Trim();
                string strTechQ = ((TextBox)Page.FindControl("txtTechQ" + strIndex)).Text.Trim();
                string strAnswer = ((TextBox)Page.FindControl("txtAnswer" + strIndex)).Text.Trim();


                HtmlInputHidden hidStartTime = (HtmlInputHidden)Page.FindControl("hidStartTime" + strIndex);
                string strStartTime = hidStartTime.Value.Trim();
                HtmlInputHidden hidEndTime = (HtmlInputHidden)Page.FindControl("hidEndTime" + strIndex);
                string strEndTime = hidEndTime.Value.Trim();
                RadioButtonList radException = (RadioButtonList)Page.FindControl("radException" + strIndex);

                string strException = radException.SelectedValue.Trim();
                RadioButtonList radClose = (RadioButtonList)Page.FindControl("radClose" + strIndex);
                string strClose = radClose.SelectedValue.Trim();
                string strSource = ((DropDownList)Page.FindControl("drpSource" + strIndex)).SelectedValue.Trim();
                string strSourceNote = "";
                //GGTS/其它選項必塡資料
                if (strSource.Equals("1")
                    || strSource.Equals("6"))
                {
                    strSourceNote = ((TextBox)Page.FindControl("txtSource" + strIndex)).Text.Trim();

                }
                string strType1 = ((HtmlInputHidden)Page.FindControl("hidType" + strIndex)).Value.Trim();
                string strType2 = "";
                if (strType1.ToUpper().Equals("OTHERS"))
                {
                    strType2 = ((DropDownList)Page.FindControl("drpSubType" + strIndex)).SelectedValue.Trim();
                    //資料檢核必填欄位
                    if (string.IsNullOrEmpty(strType2))
                    {
                        Public_Helper.CallJavascriptEvent(this, "valrType2", "PrivateAlertMsg('" + ViewState["valrType2"].ToString() + "');", true);
                        return;
                    }
                }
                #endregion

                //主檔資料
                string strTimeCountring = "";
                string strOutBoundID = this.txtOutBoundID.Text.Trim();
                string strWeek = this.txtWeek.Text.Trim();
                string strUserID = hidUserID.Value.Trim();
                //string strTel1 = this.txtTel1.Text.Trim();
                //string strTel2 = this.txtTel2.Text.Trim();
                string strCountry = this.txtCountry.Text.Trim();
                //**************************************
                //2012-04-12 個資法修正 統一整理
                string strTel1 = strPrivacy(txtTel1.Text, HidTel1.Value);
                string strTel2 = strPrivacy(txtTel2.Text, HidTel2.Value);
                string strEmail = strPrivacy(txtEmail.Text, HidEmail.Value);
                string strAddress1 = strPrivacy(this.txtAddress1.Text, this.HidAddress1.Value);
                string strAddress2 = strPrivacy(this.txtAddress2.Text, this.HidAddress2.Value);
                //**************************************


                //********20101224 加省份**********//
                objOutBoundEntity.CITY_NAME = "";
                objOutBoundEntity.CITY_ID = "";
                //20120809_Claire修改省份為下拉
                //string strCityID = this.txtChinaCity.Text.Trim();
                string strCityID = this.drpChinaCity.SelectedValue.Trim();
                string strCityName = "";
                if (strCountry.Equals("086"))
                {
                    if (!string.IsNullOrEmpty(strCityID))
                    {
                        if (strCityID.Equals("其他"))
                        {
                            strCityName = this.txtChinaCityO.Text.Trim();
                            if (string.IsNullOrEmpty(strCityName))
                            {
                                //省份選其他，名稱不能為空
                                Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrCityName"].ToString() + "');", true);
                                return;
                            }
                            objOutBoundEntity.CITY_NAME = strCityName;
                            //20120809_Claire新增
                            objOutBoundEntity.CITY_ID = strCityID;
                        }
                        else
                        {
                            //20120809_Claire修改
                            //objOutBoundEntity.CITY_ID = objOutBoundOpe.GetCity_ID(strCityID.Trim());
                            objOutBoundEntity.CITY_ID = strCityID;
                            if (string.IsNullOrEmpty(objOutBoundEntity.CITY_ID))
                            {
                                //省份輸入有誤
                                Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxCity"].ToString() + "');", true);
                                return;
                            }
                        }
                    }
                    else
                    {
                        //省份沒有輸入
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrCity"].ToString() + "');", true);
                        return;
                    }
                }
                //**********************************//
                string strSex = "";
                if (radSir.Checked)
                {
                    strSex = "M";
                }
                else if (radMis.Checked)
                {
                    strSex = "F";
                }
                string str4Num = this.txt4Num.Text.Trim();
                //string strEmail = this.txtEmail.Text.Trim();
                string strUserType = drpUserType.SelectedValue.Trim();
                //string strAddress1 = this.txtAddress1.Text.Trim();
                //string strAddress2 = this.txtAddress2.Text.Trim();
                string strCustName = this.txtUserInfo.Text.Trim();

                //主檔資料未新增,檢核主檔的資料是否正確
                if (string.IsNullOrEmpty(strOutBoundID))
                {
                    //必填欄位
                    //客戶名字資料不能為空
                    if (string.IsNullOrEmpty(strCustName))
                    {
                        Public_Helper.CallJavascriptEvent(this, "valrUserInfo", "PrivateAlertMsg('" + ViewState["valrUserInfo"].ToString() + "');", true);
                        return;
                    }
                    if (string.IsNullOrEmpty(strSex))
                    {
                        Public_Helper.CallJavascriptEvent(this, "valrSex", "PrivateAlertMsg('" + ViewState["valrSex"].ToString() + "');", true);
                        return;
                    }
                    //限制字符"|"線的輸入
                    //Check_Helper checkHelper = new Check_Helper();
                    if (Check_Helper.IsBadStr(strCustName))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxUserInfo"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strTel1))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxTel1"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strTel2))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxTel2"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(str4Num))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valx4Num"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strEmail))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxEmail"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strAddress1))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxAddress1"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strAddress2))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxAddress2"].ToString() + "');", true);
                        return;
                    }
                    //20101224 加省份
                    if (Check_Helper.IsBadStr(strCityName))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxcity"].ToString() + "');", true);

                        return;
                    }

                    //長度檢核
                    #region 地址長度檢核
                    //if (150 < strAddress1.Length)
                    //{
                    //    Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxLenAddress1"].ToString() + "');", true);

                    //    return;
                    //}

                    //if (150 < strAddress2.Length)
                    //{
                    //    Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxLenAddress2"].ToString() + "');", true);

                    //    return;
                    //}

                    //20120622_Claire_修改地址長度不得少於5字元且不得超過150字元(as_resource修改)
                    if (!string.IsNullOrEmpty(strAddress1))
                    {
                        if ((strAddress1.Length < 5) || (150 < strAddress1.Length))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxLenAddress1"].ToString() + "');", true);

                            return;
                        }
                    }
                    if (!string.IsNullOrEmpty(strAddress2))
                    {
                        if ((strAddress2.Length < 5) || (150 < strAddress2.Length))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxLenAddress2"].ToString() + "');", true);

                            return;
                        }
                    }
                    #endregion
                    //20101224 加省份
                    if (50 < strCityName.Length)
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxLenCityName"].ToString() + "');", true);

                        return;
                    }

                    //20120622_Claire_電話長度判斷
                    #region 電話長度檢核
                    if (!string.IsNullOrEmpty(strTel1))
                    {
                        if (strTel1.Length < 5)
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxLenTel1"].ToString() + "');", true);
                            return;
                        }
                    }
                    if (!string.IsNullOrEmpty(strTel2))
                    {
                        if (strTel2.Length < 5)
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxLenTel2"].ToString() + "');", true);
                            return;
                        }
                    }
                    #endregion

                    //20120927_Claire增加E-MAIL卡關EmailCheck
                    #region E-MAIL格式卡關
                    if (!string.IsNullOrEmpty(strEmail))
                    {
                        string strMailCheck = objIssueOperate.EmailCheck(strEmail);
                        if (!string.IsNullOrEmpty(strMailCheck))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxEmailFormat"].ToString() + "');", true);

                            return;
                        }
                    }
                    #endregion


                }

                #region OutBoundRecord檢核
                //OutBoundRecord 資料未存在,新增OutBoundRecord資料并進行檢核
                if (string.IsNullOrEmpty(strOutBoundSn))
                {
                    //資料檢核 必填欄位
                    if (string.IsNullOrEmpty(strModelID))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrModel"].ToString() + "');", true);
                        return;
                    }
                    //選擇爲異常  Tech可爲空
                    if (!strException.Equals("Y"))
                    {
                        if (string.IsNullOrEmpty(strTech))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrTech"].ToString() + "');", true);
                            return;
                        }
                    }
                    if (string.IsNullOrEmpty(strTechQ))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrTechQ"].ToString() + "');", true);
                        return;
                    }
                    if (string.IsNullOrEmpty(strAnswer))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrAnswer"].ToString() + "');", true);
                        return;
                    }


                    //保留字符”|“輸入的檢核

                    if (Check_Helper.IsBadStr(strModelID))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxModelID"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strRmaNo))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxRmaNo"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strRelateID))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxRelateID"].ToString() + "');", true);
                        return;
                    }

                    if (Check_Helper.IsBadStr(strTechQ))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxTechQ"].ToString() + "');", true);
                        return;
                    }
                    if (Check_Helper.IsBadStr(strAnswer))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxAnswer"].ToString() + "');", true);
                        return;
                    }

                    //有點選了Test/Call close狀態不選為Y
                    if (!string.IsNullOrEmpty(strAction))
                    {
                        if (strClose.Equals("Y"))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxAction"].ToString() + "');", true);
                            return;
                        }
                    }
                    //2009/7/23加入資料來源不能爲空
                    if (string.IsNullOrEmpty(strSource))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrSource"].ToString() + "');", true);

                        return;
                    }
                    //GGTS/其它選項必塡資料
                    if (strSource.Equals("1")
                        || strSource.Equals("6"))
                    {
                        if (string.IsNullOrEmpty(strSourceNote))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valrSourceNote"].ToString() + "');", true);

                            return;
                        }
                    }

                    //檢核相關的單號資料庫是否存在
                    if (!objIssueOperate.CheckRelateID(strRelateID))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxRelateID1"].ToString() + "');", true);

                        return;
                    }
                }
                #endregion

                //Alarm資料檢核

                if (Check_Helper.IsBadStr(strSubject))
                {
                    Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxSubject"].ToString() + "');", true);
                    return;
                }

                //如果輪入了TO/CC Email查詢Email是否正確
                if (!string.IsNullOrEmpty(strTO))
                {
                    if (strTO.IndexOf("@") > -1)
                    {
                        if (!Check_Helper.ValidateUserInput(strTO, 16))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxTo"].ToString() + "');", true);
                            return;
                        }
                    }
                    else
                    {
                        //判斷系統輸入的名字是否存在已系統中
                        string strTO1 = objIssueOperate.GetUserID(strTO);
                        if (string.IsNullOrEmpty(strTO1))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + strTO + " " + ViewState["valxTo1"].ToString() + "');", true);
                            return;
                        }
                        strTO = strTO1;
                    }
                }

                //查找塡入的To接收者資料是否已經存在
                strTO = CheckExist(strTO, strAttenTo);
                if (!string.IsNullOrEmpty(strCC))
                {
                    if (strCC.IndexOf("@") > -1)
                    {
                        if (!Check_Helper.ValidateUserInput(strCC, 16))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["valxCC"].ToString() + "');", true);
                            return;
                        }
                    }
                    else
                    {
                        //判斷系統輸入的名字是否存在已系統中
                        string strCC1 = objIssueOperate.GetUserID(strCC);
                        if (string.IsNullOrEmpty(strCC1))
                        {
                            Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + strCC + " " + ViewState["valxTo1"].ToString() + "');", true);
                            return;
                        }
                        strCC = strCC1;
                    }
                }
                strCC = CheckExist(strCC, strAttenCC);
                //新增時,自動紀錄CC欄位爲登錄帳號
                //有客到,默認CC 爲登錄帳號
                if (string.IsNullOrEmpty(strAlarmID))
                {
                    if (strAlarmType.Equals("G"))
                    {
                        if (strCC.IndexOf(Session["g_UserID"].ToString()) == -1)
                        {
                            if (string.IsNullOrEmpty(strCC))
                            {
                                strCC = Session["g_UserID"].ToString();
                            }
                            else
                            {

                                strCC += "," + Session["g_UserID"].ToString();
                            }
                        }
                    }
                }
                strCC = CheckExist(strCC, strAttenCC);
                #region 20120829_Claire增加收件人卡關：IS_LOGIN='Y'才可發信
                //20120829_Claire增加收件人卡關：IS_LOGIN='Y'才可發信
                if ((!string.IsNullOrEmpty(objOutBoundEntity.ATTEN_TO)) || (!string.IsNullOrEmpty(objOutBoundEntity.ATTEN_CC)))
                {

                    string strAddressee = objOutBoundOpe.CheckAddressee(objOutBoundEntity);
                    if (!string.IsNullOrEmpty(strAddressee))
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["AddresseeError"].ToString() + " : " + strAddressee + "');", true);
                        return;
                    }
                }
                #endregion


                #region 新增資料覆值

                objOutBoundEntity.OUTBOUND_ID = strOutBoundID;
                objOutBoundEntity.TIME_COUNTING = strTimeCountring;
                objOutBoundEntity.WEEK = strWeek;
                objOutBoundEntity.USER_ID = strUserID;
                objOutBoundEntity.TEL_NUM_1 = strTel1;
                objOutBoundEntity.TEL_NUM_2 = strTel2;
                objOutBoundEntity.COUNTRY = strCountry;
                objOutBoundEntity.CREATE_ID = Session["g_UserID"].ToString().Trim();
                objOutBoundEntity.CUST_ID = strUserID;
                objOutBoundEntity.CUST_INVOICE_ADDRESS = strAddress2;
                objOutBoundEntity.CUST_SHIPPING_ADDRESS = strAddress1;
                objOutBoundEntity.SEX = strSex;
                objOutBoundEntity.FOUTH_NUM = str4Num;
                objOutBoundEntity.EMAIL = strEmail;
                objOutBoundEntity.USER_TYPE = strUserType;
                objOutBoundEntity.CUST_SHORT_NAME = strCustName;
                objOutBoundEntity.OUTBOUND_SN = strOutBoundSn;
                objOutBoundEntity.PRODUCT_LINE = strProductLine;
                objOutBoundEntity.MODEL_ID = strModelID;

                objOutBoundEntity.ACTION = strAction;
                objOutBoundEntity.TECH = strTech;
                objOutBoundEntity.RMA_ORDER_NO = strRmaNo;
                objOutBoundEntity.EXCEPTION = strException;
                objOutBoundEntity.CLOSE = strClose;
                objOutBoundEntity.RELATE_ID = strRelateID;

                objOutBoundEntity.CLASS_VALUE = ((HtmlInputHidden)Page.FindControl("hidClassValue" + strIndex)).Value.Trim();
                objOutBoundEntity.PRODUCT_ID = ((HtmlInputHidden)Page.FindControl("hidProductID" + strIndex)).Value.Trim();
                objOutBoundEntity.PRODUCT_VERSION = strProductVersion;
                objOutBoundEntity.IS_SUBMIT = "N";
                objOutBoundEntity.SOURCE = strSource;
                //記錄用點點擊開如新增時間
                objOutBoundEntity.ACTION_START_DATE = strStartTime;
                objOutBoundEntity.ACTION_END_DATE = strEndTime;
                objOutBoundEntity.TECH_Q = strTechQ;
                objOutBoundEntity.ANSWER = strAnswer;
                objOutBoundEntity.SEQ = strIndex;
                objOutBoundEntity.ALARM_DATE = strDate;
                objOutBoundEntity.ALARM_TYPE = strAlarmType;
                objOutBoundEntity.TYPE1 = strType1;
                objOutBoundEntity.TYPE2 = strType2;
                objOutBoundEntity.SOURCE_NOTE = strSourceNote;
                if (!string.IsNullOrEmpty(strCC))
                {
                    if (!string.IsNullOrEmpty(strAttenCC))
                    {
                        objOutBoundEntity.ATTEN_CC = strCC + "," + strAttenCC;
                    }
                    else
                        objOutBoundEntity.ATTEN_CC = strCC;
                }
                else
                {
                    objOutBoundEntity.ATTEN_CC = strAttenCC;
                }

                if (!string.IsNullOrEmpty(strTO))
                {
                    if (!string.IsNullOrEmpty(strAttenTo))
                    {
                        objOutBoundEntity.ATTEN_TO = strTO + "," + strAttenTo;
                    }
                    else
                    {
                        objOutBoundEntity.ATTEN_TO = strTO;
                    }
                }
                else
                {
                    objOutBoundEntity.ATTEN_TO = strAttenTo;
                }
                objOutBoundEntity.ALARM_ID = strAlarmID;
                objOutBoundEntity.SUBJECT = strSubject;
                #endregion

                //strAlarmID 不爲空修改,否則爲新增 
                if (string.IsNullOrEmpty(strAlarmID))
                {
                    string strOutBoundID1 = "";
                    string strUserID1 = "";
                    string strOutBoundSn1 = "";
                    string strAlarmID1 = "";


                    if (objOutBoundOpe.AddActionAlarm(objOutBoundEntity, out strOutBoundID1, out strUserID1, out strOutBoundSn1, out strAlarmID1, AddAlarm))
                    {
                        //模式切換的到只讀模式 hidAlarmStatus保存着是新增模式,還是修改模式 值爲ADD新增,MODIFY修改,空值爲只讀模式
                        HtmlInputHidden hidAlarmStatus = (HtmlInputHidden)Page.FindControl("hidAlarmStatus" + strIndex);
                        hidAlarmStatus.Value = "";
                        this.txtOutBoundID.Text = strOutBoundID1;
                        this.hidUserID.Value = strUserID1;
                        hidOutBoundSn.Value = strOutBoundSn1;
                        hidAlarmID.Value = strAlarmID1;

                        if (string.IsNullOrEmpty(strOutBoundSn))
                        {
                            SetOutBoundRecordStyle(strIndex);
                        }
                        strSubject = strSubject.Replace("【OUTBOUND_ID】", "【" + this.txtOutBoundID.Text + "】");

                        ((HtmlInputHidden)Page.FindControl("hidSelectedItem" + strIndex)).Value = strAlarmID1 + "|"
                            + objOutBoundEntity.ALARM_DATE + "|" + objOutBoundEntity.ATTEN_TO
                            + "|" + objIssueOperate.GetAttenName(objOutBoundEntity.ATTEN_TO) + "|" + objOutBoundEntity.ATTEN_CC +
                            "|" + objIssueOperate.GetAttenName(objOutBoundEntity.ATTEN_CC) + "|" + strAlarmType + "|" + strSubject;
                        ((TextBox)Page.FindControl("txtSubject" + strIndex)).Text = strSubject;

                        //點選爲叫修/填寫叫修資料
                        if (radAlarmType.SelectedIndex == 3)
                        {
                            hidAlarmStatus.Value = "MODIFY";
                            Public_Helper.CallJavascriptEvent(this, "Open", "RePageStatus();OpenOutBoundRepair('" + strIndex + "','" + Check_Helper.ConvertFormat(ViewState["vsAgentID"].ToString(), 3) + "','" + Check_Helper.ConvertFormat(strAlarmID1, 3) + "');", true);
                        }
                        else
                        {
                            //有客到
                            if (radAlarmType.SelectedIndex == 4)
                            {
                                //2012-01-31 同步Issue
                                Session["UserTel"] = "";
                                Session["UserTel"] = this.txtTel1.Text.ToString().Trim();

                                //Public_Helper.CallJavascriptEvent(this, "Open", "RePageStatus();OpenOutBoundGuest('" + strIndex + "','" + Check_Helper.ConvertFormat(ViewState["vsAgentID"].ToString(), 3) + "','" + Check_Helper.ConvertFormat(strAlarmID1, 3) + "');", true);
                                Public_Helper.CallJavascriptEvent(this, "Open", "RePageStatus();OpenOutBoundGuest('" + strIndex + "','" + Check_Helper.ConvertFormat(ViewState["vsAgentID"].ToString(), 3) + "','" + Check_Helper.ConvertFormat(strAlarmID1, 3) + "','" + Check_Helper.ConvertFormat(strUserInfo, 3) + "');", true);

                            }
                            else
                                Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["AddHeaderSuccess"].ToString() + "');", true);
                        }

                        //信息Html顯示
                        switch (strIndex)
                        {
                            case "1":
                                divTechQ1.InnerHtml = objOutBoundEntity.TECH_Q;
                                divAnswer1.InnerHtml = objOutBoundEntity.ANSWER;
                                break;
                            case "2":
                                divTechQ2.InnerHtml = objOutBoundEntity.TECH_Q;
                                divAnswer2.InnerHtml = objOutBoundEntity.ANSWER;
                                break;
                            case "3":
                                divTechQ3.InnerHtml = objOutBoundEntity.TECH_Q;
                                divAnswer3.InnerHtml = objOutBoundEntity.ANSWER;
                                break;
                            case "4":
                                divTechQ4.InnerHtml = objOutBoundEntity.TECH_Q;
                                divAnswer4.InnerHtml = objOutBoundEntity.ANSWER;
                                break;
                            case "5":
                                divTechQ5.InnerHtml = objOutBoundEntity.TECH_Q;
                                divAnswer5.InnerHtml = objOutBoundEntity.ANSWER;
                                break;
                        }

                    }
                    else
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["AddHeaderFailure"].ToString() + "');", true);
                        return;
                    }
                }
                else
                {

                    if (objOutBoundOpe.ModifyActionAlarm(objOutBoundEntity))
                    {
                        //模式切換的到只讀模式 hidAlarmStatus保存着是新增模式,還是修改模式 值爲ADD新增,MODIFY修改,空值爲只讀模式
                        HtmlInputHidden hidAlarmStatus = (HtmlInputHidden)Page.FindControl("hidAlarmStatus" + strIndex);
                        hidAlarmStatus.Value = "";

                        //取得新增時AlarmType狀態
                        string strSelectedItem = ((HtmlInputHidden)Page.FindControl("hidSelectedItem" + strIndex)).Value;
                        string[] arrSelectedItem = strSelectedItem.Split('|');
                        //arrSelectedItem[6]爲AlarmType
                        ((HtmlInputHidden)Page.FindControl("hidSelectedItem" + strIndex)).Value = strAlarmID + "|"
                          + objOutBoundEntity.ALARM_DATE + "|" + objOutBoundEntity.ATTEN_TO
                          + "|" + objIssueOperate.GetAttenName(objOutBoundEntity.ATTEN_TO) + "|" + objOutBoundEntity.ATTEN_CC +
                          "|" + objIssueOperate.GetAttenName(objOutBoundEntity.ATTEN_CC) + "|" + arrSelectedItem[6] + "|" + strSubject;

                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["ModifyHeaderSuccess"].ToString() + "');", true);
                    }
                    else
                    {
                        Public_Helper.CallJavascriptEvent(this, "alert", "PrivateAlertMsg('" + ViewState["ModifyHeaderFailure"].ToString() + "');", true);
                        return;
                    }
                }



                ViewState["vsLoadStatus"] = "1";
                ((TextBox)Page.FindControl("txtTo" + strIndex)).Text = "";
                ((TextBox)Page.FindControl("txtCC" + strIndex)).Text = "";
                App_WebForm_Web_Common_SmartPageControl objSmartPageCountrol = (App_WebForm_Web_Common_SmartPageControl)Page.FindControl("objPageControlAlarm" + strIndex);
                objSmartPageCountrol.BindData1();
            }
            catch (Exception Exp)
            {

                throw new Exception(Exp.Message);
            }
        }

    }
}
