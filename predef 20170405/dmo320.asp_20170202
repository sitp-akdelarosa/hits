<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo320.asp				_/
'_/	Function	:事前実搬入入力画面(表示)		_/
'_/	Date		:2003/05/29				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/08/07	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/	Modify		:	2010/05/07	コンテナ情報と一致しない
'						ということでテアウェイトが
'						100以下の場合は100倍する	_/
'_/	Modify		:	20170118 T.Okui 設定温度を各社ビューから取得_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
	Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->
<!--#include File="../download/download.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->

<%
'セッションの有効性をチェック
  CheckLoginH
  WriteLogH "b402", "実搬入事前情報入力","11",""

'データを取得
  dim SakuNo,BookNo,CONnum
  dim User,TruckerSubName,TruckerName
  
  dim file1,gerrmsg
  
  SakuNo = Request("SakuNo")
  BookNo = Request("BookNo")
  CONnum = Request("CONnum")
  '2010/05/07 Add-S Tanaka
  dim TareWeight
  '2010/05/07 Add-E Tanaka

'データをDBより検索
  dim CMPcd,HedId,Hmon,Hday
  dim UpFlag,TruckerFlag,WkCNo,compFlag
  dim Comment1,Comment2,Comment3	'C-002
  
  User   = Session.Contents("userid")
  
'エラートラップ開始
'  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

  StrSQL = "SELECT ITC.WkContrlNo, ITC.RegisterCode, ITC.TruckerSubCode1, ITC.TruckerSubCode2, "&_
           "ITC.TruckerSubCode3, ITC.TruckerSubCode4, ITC.HeadID, ITC.WorkDate, ITC.WorkCompleteDate, "&_
           "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
           "ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
           "ITC.TruckerSubName1, ITC.TruckerSubName2, ITC.TruckerSubName3, ITC.TruckerSubName4, ITC.TruckerSubName5, "&_
           "T1.Trucked AS Trucked1, T2.Trucked AS Trucked2, T3.Trucked AS Trucked3, T4.Trucked AS Trucked4 "&_
           "FROM hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
           "LEFT JOIN mTrucker T1 ON (ITC.TruckerSubCode1 = T1.HeadCompanyCode) "&_
           "LEFT JOIN mTrucker T2 ON (ITC.TruckerSubCode2 = T2.HeadCompanyCode) "&_
           "LEFT JOIN mTrucker T3 ON (ITC.TruckerSubCode3 = T3.HeadCompanyCode) "&_
           "LEFT JOIN mTrucker T4 ON (ITC.TruckerSubCode4 = T4.HeadCompanyCode) "&_
           "WHERE ITC.ContNo='"&CONnum&"' AND ITC.WkNo='"& SakuNo &"' AND ITC.WkType='3' AND ITC.Process='R'"

'C-002 ADD This Line :            "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b402","11","実搬入：データ取得","102","SQL:<BR>"&StrSQL
  end if
  WkCNo = Trim(ObjRS("WkContrlNo"))
  CMPcd    = Array("","","","","")
  CMPcd(0) = Trim(ObjRS("RegisterCode"))
  CMPcd(1) = Trim(ObjRS("TruckerSubCode1"))
  CMPcd(2) = Trim(ObjRS("TruckerSubCode2"))
  CMPcd(3) = Trim(ObjRS("TruckerSubCode3"))
  CMPcd(4) = Trim(ObjRS("TruckerSubCode4"))

'ログインユーザによって会社コード表示制御
      chengeCompCd CMPcd, UpFlag
      '2009/07/24 M.Marquez Add-S
      if UpFlag="" then 
        UpFlag=1
      end if
      '2009/07/24 M.Marquez Add-S
      If UpFlag <> 5 Then
        TruckerFlag= Trim(ObjRS("TruckerFlag"&UpFlag))
      Else
        TruckerFlag = 0
      End If
'ログインユーザによってヘッドID表示制御
    HedId  = Trim(ObjRS("HeadID"))
    IF TruckerFlag = 1 Then 
      HedId  = "*****"
    End If
'2009/08/04 Tanaka Upd-S
'ログインユーザによって担当者名称を判断
'	Select Case Trim(User)
'		Case Trim(ObjRS("RegisterCode"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName1"))
'		Case Trim(ObjRS("Trucked1"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName2"))
'		Case Trim(ObjRS("Trucked2"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName3"))
'		Case Trim(ObjRS("Trucked3"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName4"))
'		Case Trim(ObjRS("Trucked4"))
'			TruckerSubName = Trim(ObjRS("TruckerSubName5"))
'		Case Else
'			TruckerSubName = ""
'	End Select 
	Select Case UCase(Trim(User))
		Case UCase(Trim(ObjRS("RegisterCode")))
			TruckerSubName = Trim(ObjRS("TruckerSubName1"))
			TruckerName = ObjRS("TruckerSubName1")
		Case UCase(Trim(ObjRS("Trucked1")))
			TruckerSubName = Trim(ObjRS("TruckerSubName1"))
			TruckerName = ObjRS("TruckerSubName2")
		Case UCase(Trim(ObjRS("Trucked2")))
			TruckerSubName = Trim(ObjRS("TruckerSubName2"))
			TruckerName = ObjRS("TruckerSubName3")
		Case UCase(Trim(ObjRS("Trucked3")))
			TruckerSubName = Trim(ObjRS("TruckerSubName3"))
			TruckerName = ObjRS("TruckerSubName4")
		Case UCase(Trim(ObjRS("Trucked4")))
			TruckerSubName = Trim(ObjRS("TruckerSubName4"))
			TruckerName = ObjRS("TruckerSubName5")
		Case Else
			TruckerSubName = ""
	End Select 
	
'2009/08/04 Tanaka Upd-E
'搬入予定日
  dim TmpA
  If ObjRS("WorkDate") = "1900/01/01" Or IsNull(ObjRS("WorkDate")) Then	'日付がNullであった場合
    Hmon   = Null
    Hday   = Null
  Else
    TmpA   = Split(ObjRS("WorkDate"), "/", -1, 1)
    Hmon   = TmpA(1)
    Hday   = TmpA(2)
  End If
  compFlag  = isNull(ObjRS("WorkCompleteDate"))
  Comment1=Trim(ObjRS("Comment1"))
  Comment2=Trim(ObjRS("Comment2"))
  Comment3=Trim(ObjRS("Comment3"))
  ObjRS.close
  
'20170119 T.Okui Upd(BookingからSenderの値を取得)
'2016/10/14 H.Yoshikawa Upd(VslCodeはBookingの値を取得、VoyCtrlを追加)
'2016/11/02 H.Yoshikawa Upd(PlaceDel,LPortはCYVaninfoの値を取得、それぞれの名称も取得)
  StrSQL = "SELECT CYV.ShipLine, CYV.VslName, CYV.Voyage, CYV.DPort, CYV.DelivPlace, CYV.ContSize, CYV.ContType, "&_
           "CYV.ContHeight, CYV.Material, CYV.TareWeight, CYV.CustOK, CYV.SealNo, CYV.ContWeight, CYV.ReceiveFrom, "&_
           "CYV.CustClear, CYV.OvHeight, CYV.OvWidthL, CYV.OvWidthR, CYV.OvLengthF, CYV.OvLengthA, CYV.Operator, "&_
           "EXC.RHO, CYV.SetTemp, CYV.Ventilation, CYV.IMDG1, CYV.IMDG2, CYV.IMDG3, CYV.UNNo1, CYV.UNNo2,CYV.UNNo3,"&_
           "BOK.RecTerminal, CYV.PlaceRec, CYV.LPort, "&_
           "CYV.ReportNo, CYV.AsDry, CYV.IMDG4, CYV.IMDG5, CYV.UNNo4, CYV.UNNo5, CYV.Label1, CYV.Label2, CYV.Label3, CYV.Label4, CYV.Label5, "&_
           "CYV.SubLabel1, CYV.SubLabel2, CYV.SubLabel3, CYV.SubLabel4, CYV.SubLabel5, CYV.LqFlag1, CYV.LqFlag2, CYV.LqFlag3, CYV.LqFlag4, CYV.LqFlag5, "&_
           "CYV.Solas, CYV.Consent, CYV.ContactInfo, CYV.PRShipper, CYV.PRForwarder, CYV.PRForwarderTan, CYV.PRForwarderTel, BOK.VslCode, sl.FullName as ShipLineName "&_
           ",BOK.VoyCtrl "&_
           ",mP_LP.FullName AS LPortNm, mP_DP.FullName AS DPortNm, mP_WP.FullName AS NiwataNm, mP_UP.FullName AS NiukeNm "&_
           ",BOK.Sender "&_
           "FROM (CYVanInfo AS CYV LEFT JOIN ExportCont AS EXC ON (CYV.ContNo = EXC.ContNo) AND "&_
           "(CYV.BookNo = EXC.BookNo)) LEFT JOIN Booking AS BOK ON (EXC.VslCode = BOK.VslCode) AND "&_
           "(EXC.VoyCtrl = BOK.VoyCtrl) AND (EXC.BookNo = BOK.BookNo) "&_
           "LEFT JOIN mPort AS mP_LP ON CYV.LPort = mP_LP.PortCode "&_
           "LEFT JOIN mPort AS mP_DP ON CYV.DPort = mP_DP.PortCode "&_
           "LEFT JOIN mPort AS mP_WP ON CYV.DelivPlace = mP_WP.PortCode "&_
           "LEFT JOIN mPort AS mP_UP ON CYV.PlaceRec = mP_UP.PortCode "&_
           "LEFT JOIN mShipLine AS sl ON sl.ShipLine = CYV.ShipLine "&_
          "WHERE CYV.BookNo='"& BookNo &"' AND CYV.ContNo='"& CONnum &"' AND CYV.WkNo='"& SakuNo &"' "
'20040227 Change Bok.PlaceDel->CASE WHEN mP.FullName IS Null Then Bok.PlaceRec Else mP.FullName END AS PlaceDel,
'20040227 ADD LEFT JOIN mPort AS mP ON Bok.PlaceRec = mP.PortCode
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b402","11","実搬入：データ取得","102","SQL:<BR>"&StrSQL
  end if
  dim TuSk
  If Trim(ObjRS("CustClear")) = "N" Then
    TuSk="未"
  ElseIf Trim(ObjRS("CustClear")) = "Y" Then
    TuSk="済"
  Else
    TuSk="　"
  End If
  '2010/05/07 Add-S Tanaka
  TareWeight=Trim(ObjRS("TareWeight"))
  If TareWeight<100 Then
     TareWeight=TareWeight*100
  End If
  '2010/05/07 Add-E Tanaka
  
  '20170118 T.Okui Add Start
  '設定温度を取得対応
  '設定温度、コンテナタイプ取得
    dim Operator,SetTemp,ContType
    Operator = ""
    SetTemp = ""
    ContType = ""
    
	select case gfTrim(ObjRS("Sender"))
		case "KACCS(博多港運)"
			Operator = "HKK"
		case "KACCS(上組)"
			Operator = "KAM"
		case "KACCS(ジェネック)"
			Operator = "KTC"
		case "KACCS(三菱倉庫)"
			Operator = "MLC"
		case "KACCS(日本通運)"
			Operator = "NEC"
		case "KACCS(相互運輸)"
			Operator = "SOG"
	end select
	
'	response.write StrSQL
'	response.write ObjRS("Sender")&"<br>"
'	response.write Operator
	
'	response.end
	dim VslCode, VoyCtrl
	
	VslCode = gfTrim(ObjRS("VslCode"))
	VoyCtrl = gfTrim(ObjRS("VoyCtrl"))
	if Operator <> "" then
		dim ExVoy,ObjRS3
	    ConnDBH ObjConn, ObjRS3
	    
		'ExVoyage取得
		StrSQL = "SELECT LdVoyage FROM VslSchedule "
		StrSQL = StrSQL & " WHERE VslCode='"& gfSQLEncode(VslCode) &"' "
		StrSQL = StrSQL & "   AND VoyCtrl='"& gfSQLEncode(VoyCtrl) &"' "
		ObjRS3.Open StrSQL, ObjConn
		'response.write StrSQL & "<br>"		
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS4	'DB切断
			jampErrerP "1","b402","11","実搬入：データ取得","102","SQL:<BR>"&StrSQL
		end if
		
		if not ObjRS3.EOF then
			ExVoy   = gfTrim(ObjRS3("LdVoyage"))
		end if
 		ObjRS3.Close
 		
 		
	    dim ObjRS4
	    ConnDBH ObjConn, ObjRS4

		'oBookContから設定温度取得
		StrSQL = "SELECT ob.SetTemp FROM " & Trim(Operator) & "_oBookCont ob "
		StrSQL = StrSQL & " INNER JOIN KAC_kMVessel mv on mv.VslCode = ob.VslCode "
		StrSQL = StrSQL & " INNER JOIN KAC_kSchedule kc on kc.VslCode = ob.VslCode and kc.Voyage = ob.Voyage "
		StrSQL = StrSQL & " WHERE ob.BookNo='"& gfSQLEncode(BookNo) &"' "
		StrSQL = StrSQL & "   AND mv.CallSign='"& gfSQLEncode(VslCode) &"' "
		StrSQL = StrSQL & "   AND kc.ExVoyage='"& gfSQLEncode(ExVoy) &"' "
		StrSQL = StrSQL & "   AND ob.ContNo='"& gfSQLEncode(CONnum) &"' "
		ObjRS4.Open StrSQL, ObjConn
		'response.write StrSQL & "<br>"
		'response.end
		
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS4	'DB切断
			jampErrerP "1","b402","11","実搬入：データ取得","102","SQL:<BR>"&StrSQL
		end if
		
		if not ObjRS4.EOF then
			SetTemp    = gfTrim(ObjRS4("SetTemp"))
		end if
		
		ObjRS4.Close
		
		dim ObjRS5
	    ConnDBH ObjConn, ObjRS5
	    
		'oContainerからコンテナタイプ取得
		StrSQL = "SELECT oc.ContType FROM " & Trim(Operator) & "_oContainer oc "
		StrSQL = StrSQL & " WHERE oc.ContNo='"& gfSQLEncode(CONnum) &"' "
		ObjRS5.Open StrSQL, ObjConn
		
		if err <> 0 then
			DisConnDBH ObjConn, ObjRS5	'DB切断
			jampErrerP "1","b402","11","実搬入：データ取得","102","SQL:<BR>"&StrSQL
		end if
		
		if not ObjRS5.EOF then
			ContType    = gfTrim(ObjRS5("ContType"))
		end if
			
		ObjRS5.Close
		
		
	end if
   '20170118 T.Okui Add End
   
  '2010/02/18 M.Marquez Add-A
  if Request.Form("Gamen_Mode")="R" then 
     wReportName="搬入票" 
     wReportID="dmo320" 
     wOutFileName=gfReceiveReport(BookNo,SakuNo,CONnum)
     file1	= server.mappath(gOutFileForder & wOutFileName)
	 if not gfdownloadFile(file1, wOutFileName) then
			wMsg = Replace(gerrmsg,"<br>","\n")
	 end if

  end if
  '2010/02/18 M.Marquez Add-E
  
  '2016/10/17 H.Yoshikawa Add Start 
  'ExportContが未登録の場合、Bookingから船名・次航を取得（ブッキング番号と船社を指定）
  '2017/01/23 T.Okui Upd Start 
  'dim VslCode, VoyCtrl, ObjRS2
  
  'VslCode = gfTrim(ObjRS("VslCode"))
  'VoyCtrl = gfTrim(ObjRS("VoyCtrl"))
  dim ObjRS2
  
  '2017/01/23 T.Okui Upd End 
  
  if VslCode = "" then
	ConnDBH ObjConn, ObjRS2

    StrSQL = "SELECT VslCode, VoyCtrl FROM Booking "
	StrSQL = StrSQL & " Where BookNo = '" & BookNo & "' "
	StrSQL = StrSQL & "   and ShipLine = '" & gfTrim(ObjRS("ShipLine")) & "' "
	ObjRS2.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS2	'DB切断
		jampErrerP "1","b402","11","実搬入：データ取得","102","SQL:<BR>"&StrSQL
	end if
	if not ObjRS2.EOF then
		VslCode = gfTrim(ObjRS2("VslCode"))
		VoyCtrl = gfTrim(ObjRS2("VoyCtrl"))
	end if
	ObjRS2.close
  end if
  '2016/10/17 H.Yoshikawa Add End 
  
  '2016/11/16 H.Yoshikawa Add Start
  dim CustOK
  if gfTrim(ObjRS("CustOK")) = "Y" then
  	CustOK = "Y"
  else
  	CustOK = "N"
  end if
  
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>搬入票作成情報入力(表示)</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  // 2016/10/24 H.Yoshikawa Upd Start
  //window.resizeTo(850,690);
  window.moveTo(120,10);
  window.resizeTo(1000,840);
  // 2016/10/24 H.Yoshikawa Upd End
  window.focus();
  len = target.elements.length;
  for (i=0; i<len-4; i++) target.elements[i].readOnly = true;
  bgset(target);
  if ("<%=wMsg%>"!=""){
        alert("<%=wMsg%>");
  }
  else{
      if ("<%=Request.Form("Gamen_Mode")%>"=="R"){
        if ("<%=wOutFileName%>"!=""){
            //openwinexcel("<%=wMsg%>","<%=wOutFileName%>");
            //fOpenExcel("<%=wIISFilePath%><%=wOutFileName%>");
            //parent.location.replace("<%=wIISFilePath%><%=wOutFileName%>");
        }
        document.dmo320F.Gamen_Mode.value="";
      }
  }
}

//コンテナ詳細画面
function GoConInfo(){
  target=document.dmo320F;
  target.BookNo.disabled=true;
  BookInfo(target);
  target.BookNo.disabled=false;
}
//更新画面へ
function GoReEntry(){
  target=document.dmo320F;
  target.action="./dmi320.asp";
  return true;
}
//2010-02-18 M.Marquez Add-S
//帳票出力画面へ
function GoReport(){
  target=document.dmo320F;
  target.Gamen_Mode.value="R";
  target.submit();
  return true;
}
function openwinexcel(msg,outfile){
    var w=500;
    var h=225;
    var l=0;
    var t=0;
    var target=document.dmo320F;


    if(screen.width){
        l=(screen.width-w)/2;
    }
    if(screen.availWidth){
        l=(screen.availWidth-w)/2;
    }
    if(screen.height){
        t=(screen.height-h)/2;
    }
    if(screen.availHeight){
        t=(screen.availHeight-h)/2;
    }
    var Win = window.open("/ExcelCreator/DownloadScreen.asp?Origin=0&OutFile=" + outfile + "&msg=" + msg, "", "width="+w+",height=" + h +",top="+t+",left="+l+",status=no,resizable=yes,scrollbars=no");
}

function fOpenExcel(lFileName) {
    var Excel, Book; 
    // Create the Excel application object.
    Excel = new ActiveXObject("Excel.Application"); 
    // Make Excel visible.
    Excel.Visible = true; 
    // Open work book.
    Book = Excel.Workbooks.Open(lFileName,false)
}
//2010-02-18 M.Marquez Add-E
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmo320F)">
<!-------------実搬入情報入力(表示)画面--------------------------->
<FORM name="dmo320F" method="POST">
<!--2010-02-18 M.Marquez Add-A-->
<INPUT type=hidden name="Gamen_Mode">
<!--2010-02-18 M.Marquez Add-E-->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2><B>搬入票作成情報入力(表示モード)</B></TD>
    <TD colspan=2><TABLE border=1 cellPadding=3 cellSpacing=0 align="right">
          <TR bgcolor="#f0f0f0"><TD>作業番号</TD><TD><%=SakuNo%></TD></TR>
        </TABLE>
    </TD></TR>
  <TR>
    <TD colspan=2 valign=top>
	<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD><DIV class=bgb>コンテナＮｏ．</DIV></TD>
	    <TD><INPUT type=text name="CONnum" value="<%=CONnum%>"></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ブッキングＮｏ．</DIV></TD>
	    <TD><INPUT type=text name="BookNo" value="<%=BookNo%>"></TD></TR>
	  <TR>
	    <TD><BR><DIV class=bgb>会社コード</DIV></TD>
	    <TD>登録者<BR>
	        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7>
	        <INPUT type=text name="CMPcd1" value="<%=CMPcd(1)%>" size=5>
	        <INPUT type=text name="CMPcd2" value="<%=CMPcd(2)%>" size=5>
	        <INPUT type=text name="CMPcd3" value="<%=CMPcd(3)%>" size=5>
	        <INPUT type=text name="CMPcd4" value="<%=CMPcd(4)%>" size=5></TD></TR>
	<!-- 2009/10/09 Add-S Fujiyama -->
	  <TR>
	    <TD Align=right><DIV class=bgb>指示元担当者</DIV></TD>
	    <TD>
	        <INPUT type=text name="TruckerName@" readonly = "readonly" value="<%=Trim(TruckerSubName)%>" maxlength=16>
	    </TD></TR>
	<!-- 2009/10/09 Add-S Fujiyama -->
	  <TR>
	    <TD><DIV class=bgb>ヘッドＩＤ</DIV></TD>
	    <TD><INPUT type=text name="HedId" value="<%=HedId%>"></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>搬入先</DIV></TD>
	    <TD><INPUT type=text name="HTo" value="<%=Trim(ObjRS("RecTerminal"))%>" size=30></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>搬入予定日</DIV></TD>
	    <TD><INPUT type=text name="Hmon" value="<%=Hmon%>" size=2>月
	        <INPUT type=text name="Hday" value="<%=Hday%>" size=2>日</TD></TR>
	  <TR>
	    <TD><DIV class=bgb>取扱船社</DIV></TD>
	    <TD><INPUT type=hidden name="ThkSya" value="<%=Trim(ObjRS("ShipLine"))%>" size=27>						<!-- 2016/08/17 H.Yoshikawa Upd (text→hidden) -->
	        <INPUT type=text name="ShipLineName" value="<%=gfTrim(ObjRS("ShipLineName"))%>" size=40>			<!-- 2016/08/17 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>本船名</DIV></TD>
	    <TD><INPUT type=text name="ShipN" value="<%=Trim(ObjRS("VslName"))%>">
	        <INPUT type=hidden name="ShipCode" value="<%=VslCode%>">												<!-- 2016/08/17 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>次航</DIV></TD>
	    <TD><INPUT type=text name="NextV" value="<%=Trim(ObjRS("Voyage"))%>">
	        <INPUT type=hidden name="VoyCtrl" value="<%=VoyCtrl%>">												<!-- 2016/10/14 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>荷受地</DIV></TD>
	    <TD><INPUT type=text name="NiukP" value="<%=Trim(ObjRS("PlaceRec"))%>" size=8>							<!-- 2016/11/03 H.Yoshikawa Upd(size追加, PlaceDel⇒PlaceRec) -->
	    	<INPUT type=text name="NiukeNm" value="<%=Trim(ObjRS("NiukeNm"))%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>積港</DIV></TD>
	    <TD><INPUT type=text name="TumiP" value="<%=Trim(ObjRS("LPort"))%>" size=8>								<!-- 2016/11/03 H.Yoshikawa Upd(size追加) -->
	    	<INPUT type=text name="LPortNm" value="<%=Trim(ObjRS("LPortNm"))%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>揚港</DIV></TD>
	    <TD><INPUT type=text name="AgeP" value="<%=Trim(ObjRS("DPort"))%>" size=8>								<!-- 2016/11/03 H.Yoshikawa Upd(size追加) -->
	    	<INPUT type=text name="DPortNm" value="<%=Trim(ObjRS("DPortNm"))%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>荷渡地</DIV></TD>
	    <TD><INPUT type=text name="NiwataP" value="<%=Trim(ObjRS("DelivPlace"))%>" size=8>						<!-- 2016/11/03 H.Yoshikawa Upd(size追加) -->
	    	<INPUT type=text name="NiwataNm" value="<%=Trim(ObjRS("NiwataNm"))%>" size=30>						<!-- 2016/11/03 H.Yoshikawa Add -->
	    </TD></TR>
	<!-- 2016/08/17 H.Yoshikawa Add Start -->
	  <TR>
	    <TD><DIV class=bgb>荷主</DIV></TD>
	    <TD><INPUT type=text name="Shipper" value="<%=gfTrim(ObjRS("PRShipper"))%>" size=40></TD>
	  </TR>
	  <TR><TD colspan=2>
	  	<div style="border: 1px solid black; margin:5px;">
	  	<table border=0 cellPadding=2 cellSpacing=0>
	  	<TR><TD colspan=2 style="color:red;">この欄が搬入票署名欄へ印字されます！！</TD></TR>
	  	<TR><TD><DIV class=bgb>*取扱海貨社名<BR>*（担当者）<BR>*（連絡先）</DIV></TD>
	    	<TD><INPUT type=text name="Forwarder" value="<%=gfTrim(ObjRS("PRForwarder"))%>" style="margin-bottom:2px;" size=40><BR>
	    		<INPUT type=text name="FwdStaff" value="<%=gfTrim(ObjRS("PRForwarderTan"))%>" ><BR>
	    		<INPUT type=text name="FwdTel" value="<%=gfTrim(ObjRS("PRForwarderTel"))%>" ></TD>
	  	</TR>
	  	</table>
	  	</div>
	  </TD></TR>
	<!-- 2016/08/17 H.Yoshikawa Add End   -->
	</TABLE>
	</TD>
    <TD colspan=2 valign=top>
	<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD>
	        <DIV class=bgb>サイズ、タイプ、高さ、テアウェイト</DIV></TD>
	    <TD><INPUT type=text name="CONsize" value="<%=Trim(ObjRS("ContSize"))%>" size=5>
	        <INPUT type=text name="CONtype" value="<%=Trim(ObjRS("ContType"))%>" size=5>
	        <INPUT type=text name="CONhite" value="<%=Trim(ObjRS("ContHeight"))%>" size=5>
	        <!--<INPUT type=text name="CONsitu" value="<%=Trim(ObjRS("Material"))%>" size=5> -->
	        <INPUT type=text name="CONtear" value="<%=Trim(ObjRS("TareWeight"))%>" size=7>kg
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>丸関</DIV></TD>
	    <TD><INPUT type=text name="MrSk" value="<%=CustOK%>" size=5></TD></TR>			<!-- 2016/11/16 H.Yoshikawa Upd(value:Trim(ObjRS("CustOK"))→CustOK) -->
	  <TR>
	    <TD><DIV class=bgb>シール番号</DIV></TD>
	    <TD><INPUT type=text name="SealNo" value="<%=Trim(ObjRS("SealNo"))%>"></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>コンテナグロス（テア含む）</DIV></TD>
	    <TD><INPUT type=text name="GrosW" value="<%=Trim(ObjRS("ContWeight"))%>" size=5>kg</TD></TR>
	  <!-- 2016/08/08 H.Yoshikawa Add Start -->
	  <TR>
		<TD colspan=2 style="padding-top:0px;">
			<INPUT type=hidden name="SolasChk" value="<%=gfTrim(ObjRS("Solas"))%>" >
			<INPUT type=checkbox <% if gfTrim(ObjRS("Solas")) = "1" then %>checked<% end if %> disabled>ここに入力したコンテナグロスはSOLAS条約に基づく方法で計測された数値です。
		</TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>届出番号または登録番号</DIV></TD>
	    <TD><INPUT type=text name="ReportNo" value="<%=gfTrim(ObjRS("ReportNo"))%>" size=35 ></TD></TR>
	  <!-- 2016/08/08 H.Yoshikawa Add End -->
	  <TR>
	    <TD><DIV class=bgb>搬入元</DIV></TD>
	    <TD><INPUT type=text name="HFrom" value="<%=Trim(ObjRS("ReceiveFrom"))%>" size=30></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ＲＨＯ</DIV></TD>
	    <TD><INPUT type=text name="RHO" value="<%=Trim(ObjRS("RHO"))%>" size=5></TD></TR>
	  <!-- 20170118 T.OKui Upd Start -->
	  <TR>
	    <TD><DIV class=bgb>設定温度</DIV></TD>
	    <% if Trim(ContType) = "RF" then %>
	    <TD><INPUT type=text name="SttiT" value="<%=Mid(Trim(SetTemp),1,5)%>" size=5>&nbsp;
	    <% 
	       dim TempDegree
	       TempDegree = Mid(Trim(SetTemp),6,1)
	       if gfTrim(TempDegree) <> "" then
	         TempDegree = "゜"&TempDegree 
	       end if
	    %>
	       <%=TempDegree%>
	    	<INPUT type=hidden name="TempDegree" value="<%=TempDegree %>">
	    	<!--
	    	<INPUT type=checkbox <% if gfTrim(ObjRS("AsDry")) = "1" then %>checked<% end if %> disabled>AS DRY
	    	<INPUT type=hidden name="AsDry" value="<%=gfTrim(ObjRS("AsDry"))%>" >
	    	-->
	    </TD>
	    <%else%>
	    <TD><INPUT type=text name="SttiT" value="" size=5>&nbsp;</TD>
	    <%end if%>
	    </TR>
	  <!-- 20170118 T.OKui Upd End -->
	  <TR>
	    <TD><DIV class=bgb>ＶＥＮＴ</DIV></TD>
	    <TD><INPUT type=text name="VENT" value="<%=Trim(ObjRS("Ventilation"))%>" size=5>&nbsp;%</TD></TR>
	  <TR>
	    <TD><DIV class=bgb>ＩＭＤＧ&nbsp;1～5</DIV></TD>
	    <TD>
        	<TABLE cellpadding=0 cellspacing=0><TR>
	        	<TD width="50px"><INPUT type=text name="IMDG1" value="<%=Trim(ObjRS("IMDG1"))%>" size=6 ></TD>	<!-- 2016/08/17 H.Yoshikawa Upd （size5→6) -->
	        	<TD width="50px"><INPUT type=text name="IMDG2" value="<%=Trim(ObjRS("IMDG2"))%>" size=6 ></TD>	<!-- 2016/08/17 H.Yoshikawa Upd （size5→6) -->
	        	<TD width="50px"><INPUT type=text name="IMDG3" value="<%=Trim(ObjRS("IMDG3"))%>" size=6 ></TD>	<!-- 2016/08/17 H.Yoshikawa Upd （size5→6) -->
	        	<TD width="50px"><INPUT type=text name="IMDG4" value="<%=Trim(ObjRS("IMDG4"))%>" size=6 ></TD>	<!-- 2016/08/17 H.Yoshikawa Add -->
	        	<TD width="50px"><INPUT type=text name="IMDG5" value="<%=Trim(ObjRS("IMDG5"))%>" size=6 ></TD>	<!-- 2016/08/17 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD></TR>
	  <!-- 2016/08/17 H.Yoshikawa Add Start -->
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>サブラベル&nbsp;1～5</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="Label1" value="<%=Trim(ObjRS("Label1"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label2" value="<%=Trim(ObjRS("Label2"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label3" value="<%=Trim(ObjRS("Label3"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label4" value="<%=Trim(ObjRS("Label4"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="Label5" value="<%=Trim(ObjRS("Label5"))%>" size=6 ></TD>
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <TR>
	    <TD style="padding-top:0px;"></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="SubLabel1" value="<%=Trim(ObjRS("SubLabel1"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel2" value="<%=Trim(ObjRS("SubLabel2"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel3" value="<%=Trim(ObjRS("SubLabel3"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel4" value="<%=Trim(ObjRS("SubLabel4"))%>" size=6 ></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel5" value="<%=Trim(ObjRS("SubLabel5"))%>" size=6 ></TD>
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <!-- 2016/08/17 H.Yoshikawa Add End -->
	  <TR>
	    <TD><DIV class=bgb>ＵＮ Ｎｏ.&nbsp;1～5</DIV></TD>
	    <TD>
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="UNNo1" value="<%=Trim(ObjRS("UNNo1"))%>" size=6 ></TD>					<!-- 2016/08/17 H.Yoshikawa Upd -->
		        <TD width="50px"><INPUT type=text name="UNNo2" value="<%=Trim(ObjRS("UNNo2"))%>" size=6 ></TD>					<!-- 2016/08/17 H.Yoshikawa Upd -->
		        <TD width="50px"><INPUT type=text name="UNNo3" value="<%=Trim(ObjRS("UNNo3"))%>" size=6 ></TD>					<!-- 2016/08/17 H.Yoshikawa Upd -->
		        <TD width="50px"><INPUT type=text name="UNNo4" value="<%=Trim(ObjRS("UNNo4"))%>" size=6 ></TD>					<!-- 2016/08/17 H.Yoshikawa Add -->
		        <TD width="50px"><INPUT type=text name="UNNo5" value="<%=Trim(ObjRS("UNNo5"))%>" size=6 ></TD>					<!-- 2016/08/17 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD></TR>
	  <!-- 2016/08/17 H.Yoshikawa Add Start -->
	  <TR>
	    <TD style="padding-top:0px;" align="right">少量危険品　</TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(ObjRS("LqFlag1")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(ObjRS("LqFlag2")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(ObjRS("LqFlag3")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(ObjRS("LqFlag4")) = "1" then %>checked<% end if %> disabled></TD>
		        <TD width="50px" align=center><INPUT type=checkbox <% if gfTrim(ObjRS("LqFlag5")) = "1" then %>checked<% end if %> disabled></TD>
		    	<INPUT type=hidden name="LqFlag1" value="<%=gfTrim(ObjRS("LqFlag1"))%>" >	    			<!-- 2016/08/17 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag2" value="<%=gfTrim(ObjRS("LqFlag2"))%>" >	    			<!-- 2016/08/17 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag3" value="<%=gfTrim(ObjRS("LqFlag3"))%>" >	    			<!-- 2016/08/17 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag4" value="<%=gfTrim(ObjRS("LqFlag4"))%>" >	    			<!-- 2016/08/17 H.Yoshikawa Add -->
		    	<INPUT type=hidden name="LqFlag5" value="<%=gfTrim(ObjRS("LqFlag5"))%>" >	    			<!-- 2016/08/17 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <!-- 2016/08/17 H.Yoshikawa Add End -->
	  <TR>
	    <TD colspan=2>
		  	<TABLE cellpadding=1 cellspacing=0><TR>
			    <TD width="40px"><DIV class=bgb>O/H</DIV></TD>
			    <TD><INPUT type=text name="OH"  value="<%=Trim(ObjRS("OvHeight"))%>"  size=5 ></TD>
			    <TD width="10px"></TD>
			    <TD width="40px"><DIV class=bgb>O/WL</DIV></TD>
			    <TD><INPUT type=text name="OWL" value="<%=Trim(ObjRS("OvWidthL"))%>" size=5 ></TD>
			    <TD width="10px"></TD>
			    <TD width="40px"><DIV class=bgb>O/WR</DIV></TD>
			    <TD><INPUT type=text name="OWR" value="<%=Trim(ObjRS("OvWidthR"))%>" size=5 ></TD>
			    <TD width="10px"></TD>
			    <TD width="40px"><DIV class=bgb>O/LF</DIV></TD>
			    <TD><INPUT type=text name="OLF" value="<%=Trim(ObjRS("OvLengthF"))%>" size=5 ></TD>
			    <TD width="10px"></TD>
			    <TD width="40px"><DIV class=bgb>O/LA</DIV></TD>
			    <TD><INPUT type=text name="OLA" value="<%=Trim(ObjRS("OvLengthA"))%>" size=5 >&nbsp;cm</TD>
			</TR></TABLE>
	    </TD></TR>
	  <TR>
	    <TD><DIV class=bgb>オペレータ</DIV></TD>
	    <TD><INPUT type=text name="Operator" value="<%=Trim(ObjRS("Operator"))%>"></TD></TR>
	</TABLE>
	</TD>
  </TR>
  <TR>
	<TD colspan=2 valign="TOP">
	<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	<!-- 2009/03/10 R.Shibuta Add-S -->
	  <TR>
	   <TD><DIV class=bgy>登録担当者</DIV></TD>
	   <TD><INPUT type=text name="TruckerSubName" readonly = "readonly" value="<%=Trim(TruckerName)%>" maxlength=16></TD>
	<!-- 2009/03/10 R.Shibuta Add-E -->
	  </TR>
<!-- 2016/08/18 H.Yoshikawa Add-S -->
	  <TR>
	   <TD><DIV class=bgy>登録者連絡先</DIV></TD>
	   <TD><INPUT type=text name="TruckerTel" value="<%=gfTrim(ObjRS("ContactInfo"))%>" onBlur="CheckLen(this,true,true,false)"></TD>
	  </TR>
<!-- 2016/10/11 H.Yoshikawa Del-S
	  <TR>
	    <TD><DIV class=bgb>備考１</DIV></TD>
	    <TD><INPUT type=text name="Comment1" value="<%=Comment1%>" size=73></TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>備考２</DIV></TD>
	    <TD><INPUT type=text name="Comment2" value="<%=Comment2%>" size=73></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>備考３</DIV></TD>
	    <TD><INPUT type=text name="Comment3" value="<%=Comment3%>" size=73></TD></TR>
2016/10/11 H.Yoshikawa Del-E   -->
	</TABLE>
	</TD>
  	<!-- 2016/08/18 H.Yoshikawa Add-S -->
	<TD colspan=2 valign="TOP">
	<TABLE border=0 cellpadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD align="left" valign="top">
	    　＜注意事項＞<BR>
	    　本画面の誤記・記入漏れは正常なる輸送を阻害しますので、<BR>
	    　入力済みの項目も含めて必ずご確認ください。<BR>
        　誤記・記入漏れにより発生する損害・費用・罰金等は<BR>
        　全て本画面入力者が負担し、船社(含むターミナル)は<BR>
        　責任を負いませんので、予めご承知おき願います。
	    </TD>
	  </TR>
	</TABLE>
	</TD>
  </TR>
<!-- 2016/08/18 H.Yoshikawa Add-S -->
  <TR>
	<TD colspan=4>
<BR>
	<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD colspan=3 align=center valign=bottom>
	    	<BR><INPUT type=checkbox <% if gfTrim(ObjRS("Consent")) = "1" then %>checked<% end if %> disabled>本画面の入力内容をゲートでの搬入票の代わりとして使用することに同意します。
	    	<BR>　※チェックがない場合は仮登録状態であり、予約受付は完了していません。
	    	<INPUT type=hidden name="AgreeChk" value="<%=gfTrim(ObjRS("Consent"))%>" >
		</TD>
	  </TR>
	<!-- 2016/08/18 H.Yoshikawa Add-E -->
	</TABLE>
	</TD>
  </TR>
  <TR>
    <TD colspan=4 align=center valign=bottom>
       <INPUT type=hidden name="SakuNo" value="<%=SakuNo%>">
       <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
<%'20030909 IF compFlag AND (UCase(Session.Contents("userid"))=CMPcd(0) Or TruckerFlag<>1) Then %>

<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
' IF UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1 Then 
  IF UCase(Session.Contents("userid"))=CMPcd(0) Or compFlag Then %>
       <INPUT type=hidden name="compFlag" value="<%=compFlag%>">
       <INPUT type=hidden name="WkCNo"     value="<%=WkCNo%>">
       <INPUT type=hidden name="TruckerFlag" value="<%=TruckerFlag%>">
       <INPUT type=hidden name="Mord" value="1">
 <%' 2009/08/04 Tanaka Add-S %>
  <INPUT type=hidden name="TruckerName" value="<%=Trim(TruckerName)%>">
 <%' 2009/08/04 Tanaka Add-E %>
       <INPUT type=submit value="更新モード" onClick="return GoReEntry()">
<%End IF%>
       <INPUT type=submit value="閉じる" onClick="window.close()">
       <INPUT type=button value="搬入票" onClick="GoReport();">
       <P>
       <INPUT type=button value="コンテナ情報" onClick="GoConInfo()"></TD></TR>
</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
<%

'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0
%>
</BODY></HTML>
