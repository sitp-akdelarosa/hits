<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										   _/
'_/	FileName	:dmi320.asp									   _/
'_/	Function	:事前実搬入入力画面(更新)					   _/
'_/	Date		:2003/05/29									   _/
'_/	Code By		:SEIKO Electric.Co 大重						   _/
'_/	Modify		:C-002	2003/08/06	備考欄追加				   _/
'_/	Modify		:3th	2003/01/31	3次変更					   _/
'_/	Modify		:20170118 T.Okui 設定温度を各社ビューから取得  _/
'_/	Modify		:	20170207 T.Okui 全体レイアウト変更		   _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->

<%
'セッションの有効性をチェック
  CheckLoginH
  If Request("Mord")="1" Then							'3Th add
    WriteLogH "b402", "実搬入事前情報入力","12",""
  End If												'3Th add

'サーバ日付の取得
 dim DayTime
 getDayTime DayTime

'データを取得
  dim CONnum,SakuNo,UpFlag,Mord,partFlg
  dim BookNo								'2016/08/08 H.Yoshikawa Add
  dim VslCode, VoyCtrl						'2016/10/14 H.Yoshikawa Add
  dim ShipLine								'2016/11/08 H.Yoshikawa Add
  SakuNo = Request("SakuNo")
  CONnum = Request("CONnum")
  UpFlag = Request("UpFlag")
  Mord   = Request("Mord")
  BookNo = Request("BookNo")				'2016/08/08 H.Yoshikawa Add
  VslCode = Request("ShipCode")				'2016/10/14 H.Yoshikawa Add
  VoyCtrl = Request("VoyCtrl")				'2016/10/14 H.Yoshikawa Add
  ShipLine = Request("ThkSya")				'2016/11/08 H.Yoshikawa Add

  dim CMPcd,HedId,i
  CMPcd = Array(Request("CMPcd0"),Request("CMPcd1"),Request("CMPcd2"),Request("CMPcd3"),Request("CMPcd4"))
  
'2009/03/10 R.Shibuta Add-S
  dim TruckerSubName
'2009/08/04 Upd-S Tanaka
'  TruckerSubName = Request("TruckerSubName")
'2016/08/08 H.Yoshikawa Upd Start
  'TruckerSubName = Request("TruckerName")
  TruckerSubName = Request("TruckerSubName")
'2016/08/08 H.Yoshikawa Upd End
'2009/08/04 Upd-E Tanaka
'2009/03/10 R.Shibuta Add-S

'ログインユーザによって表示を変更する
  HedId=Request("HedId")
  'response.write SakuNo & "-" & CONnum & "-" & Cstr(UpFlag) & "-" & Cstr(Mord) & "-" & Join(CMPcd,",") 
  'response.Write err.Description
  'response.end
  if UpFlag="" then UpFlag=1
  saveCompCd CMPcd, UpFlag
  
'2016/08/08 H.Yoshikawa Add Start
  dim ExcVslCode, ExcVoyage, ExcSize, ExcType, ExcHeight, ExcSetTemp
  dim ExcIMDG1, ExcIMDG2, ExcIMDG3, ExcIMDG4, ExcIMDG5
  dim ExcUNNo1, ExcUNNo2, ExcUNNo3, ExcUNNo4, ExcUNNo5
  dim TareWeight  '2017/03/06 T.Okui Add
  dim ObjConn1, ObjRS1, StrSQL1
  dim ExcVslName												'2016/10/14 H.Yoshikawa Add
  dim DGFlag													'2016/10/25 H.Yoshikawa Add
  dim Operator, ObjConnOpe, ObjRSOpe, ErrMsg					'2016/11/08 H.Yoshikawa Add

  ConnDBH ObjConn1, ObjRS1
  
'2016/11/08 H.Yoshikawa Upd Start
  ErrMsg = ""
  if VslCode = "" or VoyCtrl = "" then
  	ErrMsg = "指定したブッキングNoがシステムに登録されていません。<BR>入力の間違いがないか番号を確認してください。"
  end if
  'StrSQL1 = "SELECT CON.ContSize, CON.ContType, CON.ContHeight, EXC.* "
  StrSQL1 = "SELECT EXC.*, BOK.Sender "
'2016/11/08 H.Yoshikawa Upd Start
  StrSQL1 = StrSQL1 & ", MV.FullName AS VslName, SCD.LdVoyage "				'2016/10/14 H.Yoshikawa Add
  StrSQL1 = StrSQL1 & ", BOK.DGFlag "										'2016/10/25 H.Yoshikawa Add
  StrSQL1 = StrSQL1 & " FROM Booking AS BOK "
  StrSQL1 = StrSQL1 & " LEFT JOIN ExportCont AS EXC ON BOK.BookNo=EXC.BookNo AND BOK.VslCode=EXC.VslCode AND BOK.VoyCtrl=EXC.VoyCtrl AND EXC.ContNo='"& CONnum &"' "
  'StrSQL1 = StrSQL1 & " LEFT JOIN Container AS CON ON EXC.ContNo=CON.ContNo AND EXC.VslCode=CON.VslCode AND EXC.VoyCtrl=CON.VoyCtrl "			'2016/11/08 H.Yoshikawa Del
  StrSQL1 = StrSQL1 & " LEFT JOIN mVessel AS MV ON MV.VslCode=BOK.VslCode "
  StrSQL1 = StrSQL1 & " LEFT JOIN VslSchedule AS SCD ON SCD.VslCode=BOK.VslCode AND SCD.VoyCtrl=BOK.VoyCtrl "
  StrSQL1 = StrSQL1 & " WHERE BOK.BookNo='"& gfSQLEncode(BookNo) &"' "
  StrSQL1 = StrSQL1 & "   AND BOK.VslCode='"& gfSQLEncode(VslCode) &"' "					'2016/10/14 H.Yoshikawa Add
  StrSQL1 = StrSQL1 & "   AND BOK.VoyCtrl='"& gfSQLEncode(VoyCtrl) &"' "					'2016/10/14 H.Yoshikawa Add
  ObjRS1.Open StrSQL1, ObjConn1
  if not ObjRS1.EOF then
	'ExcVslCode = gfTrim(ObjRS1("VslCode"))									'2016/10/14 H.Yoshikawa Del
	'ExcVoyage  = gfTrim(ObjRS1("VoyCtrl"))									'2016/10/14 H.Yoshikawa Del
	ExcVslName = gfTrim(ObjRS1("VslName"))									'2016/10/14 H.Yoshikawa Add
	ExcVoyage  = gfTrim(ObjRS1("LdVoyage"))									'2016/10/14 H.Yoshikawa Add
	'2016/11/08 H.Yoshikawa Del Start
	'ExcSize    = gfTrim(ObjRS1("ContSize"))
	'ExcType    = gfTrim(ObjRS1("ContType"))
	'ExcHeight  = gfTrim(ObjRS1("ContHeight"))
	'2016/11/08 H.Yoshikawa Del End
	'ExcSetTemp = Replace(gfTrim(ObjRS1("SetTemp")), "C", "")				'2016/11/10 H.Yoshikawa Del
	ExcIMDG1   = gfTrim(ObjRS1("IMDG1"))
	ExcIMDG2   = gfTrim(ObjRS1("IMDG2"))
	ExcIMDG3   = gfTrim(ObjRS1("IMDG3"))
	ExcIMDG4   = gfTrim(ObjRS1("IMDG4"))
	ExcIMDG5   = gfTrim(ObjRS1("IMDG5"))
	ExcUNNo1   = gfTrim(ObjRS1("UNNo1"))
	ExcUNNo2   = gfTrim(ObjRS1("UNNo2"))
	ExcUNNo3   = gfTrim(ObjRS1("UNNo3"))
	ExcUNNo4   = gfTrim(ObjRS1("UNNo4"))
	ExcUNNo5   = gfTrim(ObjRS1("UNNo5"))
	DGFlag     = gfTrim(ObjRS1("DGFlag"))
'2016/11/08 H.Yoshikawa Add Start
	select case gfTrim(ObjRS1("Sender"))
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
'2016/11/08 H.Yoshikawa Add End	
  end if
  ObjRS1.Close

'2016/11/08 H.Yoshikawa Del Start
  ''2016/11/04 H.Yoshikawa Add Start
  ''サイズ、タイプ、ハイトが取れなかった場合は、Pickupに登録されているかをチェック
  'Dim ExcSizeAry(), ExcTypeAry(), ExcHeightAry()
  'Dim arycnt
  'if ExcSize = "" then
	'StrSQL1 = "SELECT ContSize, ContType, ContHeight FROM Pickup "
	'StrSQL1 = StrSQL1 & " WHERE BookNo='"& gfSQLEncode(BookNo) &"' "
	'StrSQL1 = StrSQL1 & "   AND VslCode='"& gfSQLEncode(VslCode) &"' "
	'StrSQL1 = StrSQL1 & "   AND VoyCtrl='"& gfSQLEncode(VoyCtrl) &"' "
	'ObjRS1.Open StrSQL1, ObjConn1
	'arycnt = 0
	'Redim ExcSizeAry(arycnt)
	'Redim ExcTypeAry(arycnt)
	'Redim ExcHeightAry(arycnt)
	'Do Until ObjRS1.EOF
	'	Redim Preserve ExcSizeAry(arycnt)
	'	Redim Preserve ExcTypeAry(arycnt)
	'	Redim Preserve ExcHeightAry(arycnt)
	'	ExcSizeAry(arycnt)   = gfTrim(ObjRS1("ContSize"))
	'	ExcTypeAry(arycnt)   = gfTrim(ObjRS1("ContType"))
	'	ExcHeightAry(arycnt) = gfTrim(ObjRS1("ContHeight"))
	'	arycnt = arycnt + 1
	'	ObjRS1.MoveNext
	'Loop 
	'ObjRS1.Close
  'else
	'Redim ExcSizeAry(0)
	'Redim ExcTypeAry(0)
	'Redim ExcHeightAry(0)
  'end if
  ''2016/11/04 H.Yoshikawa Add End
'2016/11/08 H.Yoshikawa Del End

'2016/11/08 H.Yoshikawa Add Start
  if Operator <> "" then
	Dim ExVoy
	
	'ExVoyage取得
	StrSQL1 = "SELECT LdVoyage FROM VslSchedule "
	StrSQL1 = StrSQL1 & " WHERE VslCode='"& gfSQLEncode(VslCode) &"' "					'2016/10/14 H.Yoshikawa Add
	StrSQL1 = StrSQL1 & "   AND VoyCtrl='"& gfSQLEncode(VoyCtrl) &"' "					'2016/10/14 H.Yoshikawa Add
	ObjRS1.Open StrSQL1, ObjConn1
	
	if not ObjRS1.EOF then
		ExVoy   = gfTrim(ObjRS1("LdVoyage"))
	end if
 	ObjRS1.Close
 	
 	'2016/11/14 H.Yoshikawa Del Start
	''KACCSの各OPEのDBに接続
	'ConnDBOpe Operator, ObjConnOpe, ObjRSOpe
 	'2016/11/14 H.Yoshikawa Del End
	
	''KACCSの船名、次航取得
	'Dim KACVsl, KACVoy
	'StrSQL1 = "SELECT sc.VslCode, sc.Voyage "
	'StrSQL1 = StrSQL1 & "  FROM kMVessel mv "
	'StrSQL1 = StrSQL1 & "  INNER JOIN kSchedule sc on sc.VslCode = mv.VslCode and sc.ExVoyage = '" & gfSQLEncode(ExVoy) &"' "
	'StrSQL1 = StrSQL1 & " WHERE mv.CallSign = '"& gfSQLEncode(VslCode) &"' "
	'ObjRSOpe.Open StrSQL1, ObjConnOpe
	'if not ObjRSOpe.EOF then
	'	KACVsl    = gfTrim(ObjRSOpe("VslCode"))
	'	KACVoy    = gfTrim(ObjRSOpe("Voyage"))
	'end if
	'ObjRSOpe.Close
	
	'2017/03/06 T.Okui Upd-S
	'oContainerからサイズ、タイプ、ハイト取得
	'StrSQL1 = "SELECT ContSize, ContType, ContHeight FROM " & Operator & "_oContainer "
	StrSQL1 = "SELECT ContSize, ContType, ContHeight, TareWeight FROM " & Operator & "_oContainer "
	StrSQL1 = StrSQL1 & " WHERE ContNo = '"& gfSQLEncode(CONnum) &"' "
	ObjRS1.Open StrSQL1, ObjConn1
	if not ObjRS1.EOF then
		ExcSize    = gfTrim(ObjRS1("ContSize"))
		ExcType    = gfTrim(ObjRS1("ContType"))
		ExcHeight  = gfTrim(ObjRS1("ContHeight"))
		TareWeight = gfTrim(ObjRS1("TareWeight"))
	end if
	ObjRS1.Close
	'2017/03/06 T.Okui Upd-E
	
	'oBookQtyに登録されているかもチェック
	Dim ExcSizeAry(), ExcTypeAry(), ExcHeightAry()
	Dim arycnt
	StrSQL1 = "SELECT ob.ContSize, ob.ContType, ob.ContHeight FROM " & Operator & "_oBookQty ob "
	StrSQL1 = StrSQL1 & " INNER JOIN KAC_kMVessel mv on mv.VslCode = ob.VslCode "
	StrSQL1 = StrSQL1 & " INNER JOIN KAC_kSchedule kc on kc.VslCode = ob.VslCode and kc.Voyage = ob.Voyage "
	StrSQL1 = StrSQL1 & " WHERE ob.BookNo='"& gfSQLEncode(BookNo) &"' "
	StrSQL1 = StrSQL1 & "   AND mv.CallSign='"& gfSQLEncode(VslCode) &"' "
	StrSQL1 = StrSQL1 & "   AND kc.ExVoyage='"& gfSQLEncode(ExVoy) &"' "
	'2016/11/30 H.Yoshikawa Upd Start
	'StrSQL1 = StrSQL1 & "   AND ob.Terminal='999' "
	StrSQL1 = StrSQL1 & "   AND (ob.Terminal='999' or ob.Terminal='998') "
	'2016/11/30 H.Yoshikawa Upd End

	ObjRS1.Open StrSQL1, ObjConn1
	arycnt = 0
	Redim ExcSizeAry(arycnt)
	Redim ExcTypeAry(arycnt)
	Redim ExcHeightAry(arycnt)
	Do Until ObjRS1.EOF
		Redim Preserve ExcSizeAry(arycnt)
		Redim Preserve ExcTypeAry(arycnt)
		Redim Preserve ExcHeightAry(arycnt)
		ExcSizeAry(arycnt)   = gfTrim(ObjRS1("ContSize"))
		ExcTypeAry(arycnt)   = gfTrim(ObjRS1("ContType"))
		ExcHeightAry(arycnt) = gfTrim(ObjRS1("ContHeight"))
		arycnt = arycnt + 1
		ObjRS1.MoveNext
	Loop 
	ObjRS1.Close
	
	'20170118 T.Okui Upd Start
	'2016/11/10 H.Yoshikawa Del Start
	'oBookContからSetTemp取得
	dim TempDegree
	ExcSetTemp = Request("SttiT")
	TempDegree = Request("TempDegree")
	
	if ExcType = "RF" and ExcSetTemp = "" then
		

		StrSQL1 = "SELECT ob.SetTemp FROM " & Operator & "_oBookCont ob "
		StrSQL1 = StrSQL1 & " INNER JOIN KAC_kMVessel mv on mv.VslCode = ob.VslCode "
		StrSQL1 = StrSQL1 & " INNER JOIN KAC_kSchedule kc on kc.VslCode = ob.VslCode and kc.Voyage = ob.Voyage "
		StrSQL1 = StrSQL1 & " WHERE ob.BookNo='"& gfSQLEncode(BookNo) &"' "
		StrSQL1 = StrSQL1 & "   AND mv.CallSign='"& gfSQLEncode(VslCode) &"' "
		StrSQL1 = StrSQL1 & "   AND kc.ExVoyage='"& gfSQLEncode(ExVoy) &"' "
		StrSQL1 = StrSQL1 & "   AND ob.ContNo='"& gfSQLEncode(CONnum) &"' "
		ObjRS1.Open StrSQL1, ObjConn1
		if not ObjRS1.EOF then
			ExcSetTemp = Mid(Trim(ObjRS1("SetTemp")),1,5)
			TempDegree = Mid(Trim(ObjRS1("SetTemp")),6,1)
			if gfTrim(TempDegree) <> "" then
   	          TempDegree = "゜"&TempDegree 
   	        end if
		end if
		ObjRS1.Close
	end if
	
	'20170118 T.Okui Upd End
	'OPEDB接続解除
	'DisConnDBH ObjConnOpe, ObjRSOpe			'2016/11/14 H.Yoshikawa Del

  else
	Redim ExcSizeAry(0)
	Redim ExcTypeAry(0)
	Redim ExcHeightAry(0)
	ErrMsg = "指定したブッキングNoがシステムに登録されていません。<BR>入力の間違いがないか番号を確認してください。"
  end if
'2016/11/08 H.Yoshikawa Add End
  
  Dim mIMDG()
  StrSQL1 = "SELECT IMDG FROM mIMDG "
  ObjRS1.Open StrSQL1, ObjConn1
  arycnt = 0
  Do Until ObjRS1.EOF
	Redim Preserve mIMDG(arycnt)
	mIMDG(arycnt) = gfTrim(ObjRS1("IMDG"))
	arycnt = arycnt + 1
    ObjRS1.MoveNext
  Loop 
  ObjRS1.Close

  'DB接続解除
  DisConnDBH ObjConn1, ObjRS1
'2016/08/08 H.Yoshikawa Add End

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<STYLE type="text/css">
DIV.bgb{
	text-align:center;
	margin-left:4px;
}
DIV.bgy{
	text-align:center;
	margin-left:4px;
}
</STYLE>
<TITLE>搬入票作成情報入力(登録)</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT src="./JS/CommonSub.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

function setParam(target){

//  setMonth(target.Hmon,"<%=Request("Hmon")%>");
//  setDate(target.Hday,"<%=Request("Hday")%>");
  check_date('<%=DayTime(0)%>','<%=DayTime(1)%>',target.Hmon,target.Hday);
<%
'コンボボックスデータ取得
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'コンテナサイズ取得＆表示
  StrSQL = "select * from mContSize ORDER BY ContSize ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContSize") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  Response.Write "  setList(target.CONsize,list,'" & Request("CONsize") & "');" & vbCrLf
  ObjRS.Close

'コンテナタイプ取得＆表示
  StrSQL = "select * from mContType ORDER BY ContType ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContType") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  Response.Write "  setList(target.CONtype,list,'" & Request("CONtype") & "');" & vbCrLf
  ObjRS.Close

'コンテナ高さ取得＆表示
  StrSQL = "select * from mContHeight ORDER BY ContHeight ASC"
  ObjRS.Open StrSQL, ObjConn
  Response.Write "  list = new Array(''"
  Do Until ObjRS.EOF
    Response.Write ",'" & ObjRS("ContHeight") & "'"
    ObjRS.MoveNext
  Loop 
  Response.Write ");" & vbCrLf
  Response.Write "  setList(target.CONhite,list,'" & Request("CONhite") & "');" & vbCrLf
  ObjRS.Close

'2016/08/01 H.Yoshikawa Delete Start
''コンテナ材質取得＆表示
'  StrSQL = "select * from mContMaterial ORDER BY ContMaterial ASC"
'  ObjRS.Open StrSQL, ObjConn
'  Response.Write "  list = new Array(''"
'  Do Until ObjRS.EOF
'    Response.Write ",'" & ObjRS("ContMaterial") & "'"
'    ObjRS.MoveNext
'  Loop 
'  Response.Write ");" & vbCrLf
'  Response.Write "  setList(target.CONsitu,list,'" & Request("CONsitu") & "');" & vbCrLf
'2016/08/01 H.Yoshikawa Delete End

'DB接続解除
  DisConnDBH ObjConn, ObjRS
%>
<%
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'  if(target.MrSk.options[1].value=="< %=Request("MrSk")% >"){
'    target.MrSk.selectedIndex=1;
'  } else if (target.MrSk.options[2].value=="< %=Request("MrSk")% >"){
'    target.MrSk.selectedIndex=2;
'  }
'  If Mord=0 Then 															2016/08/08 H.Yoshikawa Del
'    Response.Write "  target.MrSk.selectedIndex=2;"&Chr(10)				2016/08/08 H.Yoshikawa Del
'  Else																		2016/08/08 H.Yoshikawa Del
'2016/11/16 H.Yoshikawa Del Start
'    Response.Write "  if(target.MrSk.options[1].value=="""&Request("MrSk")&"""){"&Chr(10)&_
'                   "    target.MrSk.selectedIndex=1;"&Chr(10)&_
'                   "  } else if (target.MrSk.options[2].value=="""&Request("MrSk")&"""){"&Chr(10)&_
'                   "    target.MrSk.selectedIndex=2;"&Chr(10)&_
'                   "  }"&Chr(10)
'2016/11/16 H.Yoshikawa Del End
'  End If																	2016/08/08 H.Yoshikawa Del
'Chang 20050303 End
%>
//2016/08/02 H.Yoshikawa Delete Start
  //if(target.TuSk.options[1].value=="<%=Request("TuSk")%>"){
  //  target.TuSk.selectedIndex=1;
  //} else if (target.TuSk.options[2].value=="<%=Request("TuSk")%>"){
  //  target.TuSk.selectedIndex=2;
  //}
//2016/08/02 H.Yoshikawa Delete End

  Utype=<%=Session.Contents("UType")%>;
  if(Utype != 5) target.HedId.readOnly = true;
<%
'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
  If Mord=1 AND (Request("TruckerFlag")=1 OR Not Request("compFlag")) Then
    Response.Write "  allsetreadOnly(target,8);"&Chr(10)
    If Request("compFlag") Then
      '2016/08/01 H.Yoshikawa Upd Start
      'Response.Write "  target.SealNo.readOnly=false;"&Chr(10)&_
      '               "  target.GrosW.readOnly =false;"&Chr(10)&_
      '               "  target.Hmon.readOnly  =false;"&Chr(10)&_
      '               "  target.Hday.readOnly  =false;"&Chr(10)&_
      '               "  target.TuSk.readOnly  =false;"&Chr(10)&_
      '               "  target.CONsize.disabled =true;"&Chr(10)&_
      '               "  target.CONtype.disabled =true;"&Chr(10)&_
      '               "  target.CONhite.disabled =true;"&Chr(10)&_
      '               "  target.CONsitu.disabled =true;"&Chr(10)&_
      '               "  target.MrSk.disabled =true;"&Chr(10)
      '2017/03/10 T.Okui Upd-S
      Response.Write "  target.SealNo.readOnly=false;"&Chr(10)&_
                     "  target.GrosW.readOnly =false;"&Chr(10)&_
                     "  target.CONtear.readOnly =false;"&Chr(10)&_
                     "  target.Hmon.readOnly  =false;"&Chr(10)&_
                     "  target.Hday.readOnly  =false;"&Chr(10)&_
                     "  target.SolasChk.readOnly =false;"&Chr(10)&_
                     "  target.AgreeChk.readOnly =false;"&Chr(10)&_
                     "  target.SolasChk.disabled =false;"&Chr(10)&_
                     "  target.AgreeChk.disabled =false;"&Chr(10)
      '2017/03/10 T.Okui Upd-E
      '2016/08/01 H.Yoshikawa Upd End
      partFlg=1
    End If
  End If
'ADD 20050303 END
%>
//2016/10/25 H.Yoshikawa Add Start
  // 2016/11/10 H.Yoshikawa Del Start
  //ExcSetTemp = "<%=ExcSetTemp%>";
  //if(ExcSetTemp==""){
  //	target.SttiT.value="";
  //	target.SttiT.readOnly=true;
  //	target.AsDry.checked=false;
  //	target.AsDry.disabled=true;
  //}
  // 2016/11/10 H.Yoshikawa Del End

// 2016/11/03 H.Yoshikawa Del Start
//  DGFlag = "<%=DGFlag%>";
//  if(DGFlag==""){
//	  strA    = new Array();
//	  strA[0] = target.IMDG1;
//	  strA[1] = target.IMDG2;
//	  strA[2] = target.IMDG3;
//	  strA[3] = target.IMDG4;
//	  strA[4] = target.IMDG5;
//	  strA[5] = target.Label1;
//	  strA[6] = target.Label2;
//	  strA[7] = target.Label3;
//	  strA[8] = target.Label4;
//	  strA[9] = target.Label5;
//	  strA[10] = target.SubLabel1;
//	  strA[11] = target.SubLabel2;
//	  strA[12] = target.SubLabel3;
//	  strA[13] = target.SubLabel4;
//	  strA[14] = target.SubLabel5;
//	  strA[15] = target.UNNo1;
//	  strA[16] = target.UNNo2;
//	  strA[17] = target.UNNo3;
//	  strA[18] = target.UNNo4;
//	  strA[19] = target.UNNo5;
//	  for(k=0;k<strA.length;k++){
//	  	strA[k].value="";
//	  	strA[k].readOnly=true;
//	  }
//	  strA    = new Array();
//	  strA[0] = target.LqFlag1;
//	  strA[1] = target.LqFlag2;
//	  strA[2] = target.LqFlag3;
//	  strA[3] = target.LqFlag4;
//	  strA[4] = target.LqFlag5;
//	  for(k=0;k<strA.length;k++){
//	  	strA[k].checked=false;
//	  	strA[k].disabled=true;
//	  }
//  }
// 2016/11/03 H.Yoshikawa Del End

//2016/10/25 H.Yoshikawa Add End

  bgset(target);
  checkIDF(0);<%'CW-017 ADD%>
}

//コンテナ詳細画面
function GoConInfo(){
  target=document.dmi320F;
  target.BookNo.disabled=true;
  BookInfo(target);
  target.BookNo.disabled=false;
}
//登録・更新
function GoReEntry(){
  target=document.dmi320F;
  
  <% If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then%>
  if(target.way[1].checked){
    flag = confirm('回答をNoにしますか？');
    if(!flag) return false;
    target.Mord.value=2;
  }
  <% End If %>
  chengeUpper(target);			// 2016/10/17 H.Yoshikawa Add
  ret = check();
  if(ret==false){
    return;
  }
  target.action="./dmi330.asp";
  chengeUpper(target);
  target.submit();
}
//削除
function GoDell(){
<% If Mord<>"0" Then %>
  <%If Request("TruckerFlag")<>1 Then%>
  flag = confirm('削除しますか？');
  <%Else%>
  flag = confirm('指示先が受諾回答済です。\n削除する前に指示先に確認してください。\n削除しますか？');
  <%End If%>
  if(flag){
    target=document.dmi320F;
    target.action="./dmi390.asp";
    target.submit();
  }
<%End If%>
}

//入力情報チェック
function check(){
  target=document.dmi320F;
  strA    = new Array();
  strA[0] = target.CMPcd1;
  strA[1] = target.CMPcd2;
  strA[2] = target.CMPcd3;
  strA[3] = target.CMPcd4;
  strA[4] = target.HedId;
  strA[5] = target.SealNo;
  strA[6] = target.HFrom;
  for(k=0;k<strA.length;k++){
    if(strA[k].value!="" && strA[k].value!=null && strA[k].readOnly==false){
      ret = CheckEisu(strA[k].value); 
      if(ret==false){
        alert("半角英数字と半角スペース、「-」、「/」以外の文字を入力しないでください");
        strA[k].focus();
        return false;
      }
    }
  }
<% If UpFlag = 1 Then %>
  if(strA[0].value.length==0 && strA[4].value.length!=0){
    alert("指示先を自社に指定しなければヘッドIDを入力する事は出来ません");
    strA[0].focus();
    return false;
  }
<% End If %>
<% If partFlg<>1 Then 'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige %>
  // Added 2003.8.3
  if(strA[4].value != ""){
    if(strA[4].value.length != 5){
      alert("ヘッドＩＤは「ヘッド会社コード」＋「数字３桁」で入力してください。");
      strA[4].focus();
      return false;
    }else{
      if(isNaN(strA[4].value.charAt(2)) || isNaN(strA[4].value.charAt(3)) || isNaN(strA[4].value.charAt(4))){
        alert("ヘッドＩＤは「ヘッド会社コード」＋「数字３桁」で入力してください。");
        strA[4].focus();
        return false;
      }
    }
  }
  // End of Addition 2003.8.3
<% End If 'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige %>
  Num=LTrim(strA[5].value);
  if(Num.length==0){
    alert("シールNoを記入してください");
    strA[5].focus();
    return false;
  }
  Num=LTrim(strA[6].value);
  if(Num.length==0){
    alert("搬入元を記入してください");
    strA[6].focus();
    return false;
  }
  
//2016/10/11 H.Yoshikawa Del Start
//  //2016/08/02 H.Yoshikawa Add Start
//  Num=LTrim(strA[7].value);
//  if(Num.length==0){
//    alert("届出番号または登録番号を記入してください");
//    strA[7].focus();
//    return false;
//  }
//  //2016/08/02 H.Yoshikawa Add End
//2016/10/11 H.Yoshikawa Del End


 //2017/03/03 T.Okui Add-S
 Num=LTrim(target.CONtear.value);
 if(Num.length!=0 && Num.length!=4){
   alert("テアウェイトは4桁で入力してください。");
   target.CONtear.focus();
   return false;
 }
//2017/03/03 T.Okui Add-E

//2017/03/02 T.Okui Add-S
  strA[7] = target.ReportNo;			//2016/08/02 H.Yoshikawa Add

 Num=LTrim(strA[7].value);
 
//2017/04/05 H.Yoshikawa Upd Start
// ret = CheckSu(Num)
 ret = CheckEisu2(Num)
 if(!ret){
   //alert("届出番号または登録番号は数値で入力してください。");
   alert("届出番号または登録番号は半角英数字で入力してください。");
   strA[7].focus();
   return false;
 }
//2017/04/05 H.Yoshikawa Upd End

 if(Num.length!=0 && Num.length!=12){
   alert("届出番号または登録番号は12桁で入力してください。");
   strA[7].focus();
   return false;
 }
//2017/03/02 T.Okui Add-E

  strA    = new Array();
  
  //2016/08/02 H.Yoshikawa Upd Start
  //strA[0] = target.CONtear;
  //strA[1] = target.GrosW;
  //strA[2] = target.OH;
  //strA[3] = target.OWL;
  //strA[4] = target.OWR;
  //strA[5] = target.OLF;
  //strA[6] = target.OLA;
  //strM    = new Array("テアウェイト","グロスウェイト","Ｏ/Ｈ","Ｏ/ＷＬ","Ｏ/ＷＲ","Ｏ/ＬＦ","Ｏ/ＬＡ");
  //for(k=0;k<2;k++){
  strA[0] = target.GrosW;
  strA[1] = target.OH;
  strA[2] = target.OWL;
  strA[3] = target.OWR;
  strA[4] = target.OLF;
  strA[5] = target.OLA;
  //strM    = new Array("コンテナグロス","Ｏ/Ｈ","Ｏ/ＷＬ","Ｏ/ＷＲ","Ｏ/ＬＦ","Ｏ/ＬＡ");
  strM    = new Array("コンテナ総重量","Ｏ/Ｈ","Ｏ/ＷＬ","Ｏ/ＷＲ","Ｏ/ＬＦ","Ｏ/ＬＡ");
  for(k=0;k<1;k++){
  //2016/08/02 H.Yoshikawa Upd End
    Num=LTrim(strA[k].value);
    if(Num.length==0){
      alert(strM[k]+"を記入してください");
      strA[k].focus();
      return false;
    }
  }
  for(k=0;k<strA.length;k++){
    if(strA[k].value!="" && strA[k].value!=null){
      ret = CheckSu(strA[k].value); 
      if(ret==false){
        alert(strM[k]+"に半角数字以外を入力しないでください");
        strA[k].focus();
        return false;
      }
    }
  }
  strA    = new Array();
  strA[0] = target.CONsize;
  strA[1] = target.CONtype;
  strA[2] = target.CONhite;
  //2016/08/01 H.Yoshikawa Upd Start（材質削除）
  //strA[3] = target.CONsitu;
  //strA[4] = target.MrSk;
  //strA[5] = target.TuSk;
  //strM    = new Array("サイズ","タイプ","高さ","材質","丸関","通関");
  //strA[3] = target.MrSk;										// 2016/11/16 H.Yoshikawa Del
  strM    = new Array("サイズ","タイプ","高さ","丸関");
  //2016/08/01 H.Yoshikawa Upd Start（材質削除）
  for(k=0;k<strA.length;k++){
    if(strA[k].selectedIndex==0){
      alert(strM[k]+"を選択してください");
        strA[k].focus();
        return false;
    }
  }
  
 

<%' C-002 ADD START%>
//2016/10/11 H.Yoshikawa Del Start
//  strA[0] = target.Comment1;
//  strA[1] = target.Comment2;
//  strA[2] = target.Comment3;
//  for(k=0;k<3;k++){
//    if(strA[k].value!="" && strA[k].value!=null){
//      ret = CheckKin(strA[k].value); 
//      if(ret==false){
//        alert("「\"」や「\'」等の半角記号を入力しないでください");
//        strA[k].focus();
//        return false;
//      }
//    }
//    retA=getByte(strA[k].value);
//    if(retA[0]>70){
//      if(retA[2]>35){
//        alertStr="全角文字を35文字以内で入力してください。";
//      }else{
//        alertStr="全角文字を"+Math.floor((70-retA[1])/2)+"文字にするか\n";
//        alertStr=alertStr+"半角文字を"+(70-retA[2]*2)+"文字にしてください。";
//      }
//      alert("70バイト以内で入力してください。\n70バイト以内にするには"+alertStr);
//      strA[k].focus();
//      return false;
//    }
//  }
//2016/10/11 H.Yoshikawa Del End

  //2016/10/28 H.Yoshikawa Add Start 
  if(target.VENT.value!="" && target.VENT.value!=null){
    ret = CheckSu(target.VENT.value); 
    if(ret==false){
      //2017/04/04 H.Yoshikawa Upd Start
      //alert("VENTに半角数字以外を入力しないでください");
      alert("ベンチレーションに半角数字以外を入力しないでください");
      //2017/04/04 H.Yoshikawa Upd End
      target.VENT.focus();
      return false;
    }
    if(target.VENT.value > 100){
      //2017/04/04 H.Yoshikawa Upd Start
      //alert("VENTには0～100までの数値を入力してください");
      alert("ベンチレーションには0～100までの数値を入力してください");
      target.VENT.focus();
      return false;
    }
  }
  //2016/10/28 H.Yoshikawa Add End 

  //2016/08/02 H.Yoshikawa Add Start
  if(target.TruckerSubName.value.length==0){
    alert("登録担当者を記入してください。");
    target.TruckerSubName.focus();
    return false;
  }
  //2016/08/02 H.Yoshikawa Add End
  
  //2017/03/02 T.Okui Add Start
  if(numCheck(target.TruckerSubName.value)){
    alert("登録担当者には数値を入力しないでください。");
    target.TruckerSubName.focus();
    return false;
  }
  //2017/03/02 T.Okui Add End
  
  /* 2009/09/27 C.Pestano Del-S
  ret = CheckKana(target.TruckerSubName.value); 
  if(ret==false){
  	alert("半角カナ文字は入力できません");
  	target.TruckerSubName.focus();
  	return false;
  } 2009/09/27 C.Pestano Del-E
  */

<%' C-002 ADD END%>
<%' 3th ADD START%>
//日付のチェック
  if(!CheckDate('<%=DayTime(0)%>','<%=DayTime(1)%>',target.Hmon,target.Hday,0))
      return false;
<%' 3th ADD End%>
<%
'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
  If partFlg=1 Then
      '2016/08/01 H.Yoshikawa Upd Start
      'Response.Write "  target.CONsize.disabled =false;"&Chr(10)&_
      '               "  target.CONtype.disabled =false;"&Chr(10)&_
      '               "  target.CONhite.disabled =false;"&Chr(10)&_
      '               "  target.CONsitu.disabled =false;"&Chr(10)&_
      '               "  target.MrSk.disabled =false;"&Chr(10)
      Response.Write "  target.CONsize.disabled =false;"&Chr(10)&_
                     "  target.CONtype.disabled =false;"&Chr(10)&_
                     "  target.CONhite.disabled =false;"&Chr(10)
'                     "  target.MrSk.disabled =false;"&Chr(10)				'2016/11/16 H.Yoshikawa Del
      '2016/08/01 H.Yoshikawa Upd Start
  End If
'ADD 20050303 END
%>

//2016/08/02 H.Yoshikawa Add Start
  if(target.ShipN.value.length==0 || target.NextV.value.length==0 || target.ShipCode.value.length==0){
    alert("船名、次航が正しくありません。検索画面よりセットしてください。");
    target.ShipN.focus();
    return false;
  }
  // 2016/11/03 H.Yoshikawa Add Start
  if(target.NiukP.value.length==0){
    alert("荷受地を記入してください。");
    target.NiukP.focus();
    return false;
  }
  if(target.TumiP.value.length==0){
    alert("積港を記入してください。");
    target.TumiP.focus();
    return false;
  }
  if(target.AgeP.value.length==0){
    alert("揚港を記入してください。");
    target.AgeP.focus();
    return false;
  }
  if(target.NiwataP.value.length==0){
    alert("荷渡地を記入してください。");
    target.NiwataP.focus();
    return false;
  }
  // 2016/11/03 H.Yoshikawa Add End
  
  if(target.Shipper.value.length==0){
    alert("荷主を記入してください。");
    target.Shipper.focus();
    return false;
  }
  if(target.Forwarder.value.length==0){
    alert("取扱海貨・社名を記入してください。");
    target.Forwarder.focus();
    return false;
  }
  if(target.FwdStaff.value.length==0){
    alert("取扱海貨・担当者を記入してください。");
    target.FwdStaff.focus();
    return false;
  }
  //2017/03/02 T.Okui Add Start
  if(numCheck(target.FwdStaff.value)){
    alert("取扱海貨・担当者には数値を入力しないでください。");
    target.FwdStaff.focus();
    return false;
  }
  //2017/03/02 T.Okui Add End
  
  //2017/04/04 H.Yoshikawa Add Start
  if(numCheck(target.EntryName.value)){
    alert("指示元担当者には数値を入力しないでください。");
    target.EntryName.focus();
    return false;
  }
  //2017/04/04 H.Yoshikawa Add Start
  
  if(target.FwdTel.value.length==0){
    alert("海貨連絡先を記入してください。");
    target.FwdTel.focus();
    return false;
  }
  if(target.TruckerTel.value.length==0){
    alert("登録者連絡先を記入してください。");
    target.TruckerTel.focus();
    return false;
  }
  if(CheckTel(target.FwdTel.value)==false){
    alert("海貨連絡先が正しくありません。");
    target.FwdTel.focus();
    return false;
  }
  
  if(CheckTel(target.TruckerTel.value)==false){
    alert("登録者連絡先が正しくありません。");
    target.TruckerTel.focus();
    return false;
  }
  
  //2017/03/02 T.Okui Add Start
  var tmp;
  tmp = target.FwdTel.value;
  if(tmp.replace(/\-/g,'').length!=10 && tmp.replace(/\-/g,'').length!=11){
  	alert("海貨連絡先は10桁または11桁の番号で入力してください。");
    target.FwdTel.focus();
    return false;
  }
  
  tmp = target.TruckerTel.value;
  if(tmp.replace(/\-/g,'').length!=10 && tmp.replace(/\-/g,'').length!=11){
  	alert("登録者連絡先は10桁または11桁の番号で入力してください。");
    target.TruckerTel.focus();
    return false;
  }
  //2017/03/02 T.Okui Add End
  
  
  // 2016/10/20 H.Yoshikawa Add Start
  strA[0] = target.Shipper;
  strA[1] = target.Forwarder;
  strA[2] = target.FwdStaff;
  strA[3] = target.TruckerSubName;
  strA[4] = target.EntryName;         //2017.04.04 H.Yoshikawa Add
  for(k=0;k<strA.length;k++){
    //if(strA[k].value!="" && strA[k].value!=null){
    //  ret = CheckKin(strA[k].value); 
    //  if(ret==false){
    //    alert("「\"」や「\'」等の半角記号を入力しないでください");
    //    strA[k].focus();
    //    return false;
    //  }
    //}
    maxlen = strA[k].maxLength;
    maxlenZen = maxlen / 2 ;
    retA=getByte(strA[k].value);
    if(retA[0]>maxlen){
      alertStr="全角文字を" + maxlenZen + "文字以内にするか\n";
      alertStr=alertStr+"半角文字を"+maxlen+"文字以内にしてください。";
      alert(maxlen + "バイト以内で入力してください。\n" + maxlen + "バイト以内にするには"+alertStr);
      strA[k].focus();
      return false;
    }
  }
  // 2016/10/20 H.Yoshikawa Add End
  
  //2017/04/04 H.Yoshikawa Add Start
  wkstr = toHalfWidth(target.Forwarder.value);
  ret = CheckSu(wkstr); 
  if(ret==true){
    alert("取扱海貨社名には、数値のみの入力はできません。");
    target.Forwarder.focus();
    return false;
  }
  //2017/04/04 H.Yoshikawa Add End
  
  //2016/11/22 H.Yoshikawa Add Start
  errmsg = "「RHO」の項目にR（リーファー）がセットされている状態で、オーバーディメンションは入力できません。";
  if(target.RHO.value.indexOf("R") >= 0){
  	if(Number(target.OH.value) > 0){
		alert(errmsg);
		target.OH.focus();
		return false;
  	}
  	if(Number(target.OWL.value) > 0){
		alert(errmsg);
		target.OWL.focus();
		return false;
  	}
  	if(Number(target.OWR.value) > 0){
		alert(errmsg);
		target.OWR.focus();
		return false;
  	}
  	if(Number(target.OLF.value) > 0){
		alert(errmsg);
		target.OLF.focus();
		return false;
  	}
  	if(Number(target.OLA.value) > 0){
		alert(errmsg);
		target.OLA.focus();
		return false;
  	}
  }
  //2016/11/22 H.Yoshikawa Add End
 
  //2016/11/29 H.Yoshikawa Upd Start
  //if(target.GrosW.value < 2000){
  //	if(confirm("コンテナグロスが2,000kg未満ですが、登録してよろしいですか？") == false){
  //		return false;
  //	}
  //}
  //if(target.GrosW.value > 35000){
  //	if(confirm("コンテナグロスが35,000kgを超えていますが、登録してよろしいですか？") == false){
  //		return false;
  //	}
  //}
  if(target.SolasChk.checked && target.AgreeChk.checked){
    if(target.GrosW.value < 1800 || target.GrosW.value > 500000){
      //alert("本登録の場合、コンテナグロスは1,800Kg以上、500,000Kg以下の範囲で入力してください。");
      alert("本登録の場合、コンテナ総重量は1,800Kg以上、500,000Kg以下の範囲で入力してください。");
      target.GrosW.focus();
      return false;
    }
  }
  //2016/11/29 H.Yoshikawa Upd End

  var chkagree = "0";
  var chksolas = "0";
  var BookChk = "0";
  var IMDGChk = "0";			//2016/11/03 H.Yoshikawa Add
  var retValue;
  if(target.SolasChk.checked){
    chksolas = "1";
  }
  if(target.AgreeChk.checked){
    chkagree = "1";
  }
  if(BookingCheck() == false){
  	BookChk = "1";
  }
  //2016/11/03 H.Yoshikawa Add Start
  if(IMDGCheck() == false){
  	IMDGChk = "1";
  }
  //2016/11/03 H.Yoshikawa Add End
  
  if(chkagree == "1" && chksolas == "1" && BookChk == "0" && IMDGChk == "0"){			// 2016/11/03 H.Yoshikawa Upd (IMDGChk追加)
  	target.kariflag.value = "1";
  }else{
  	target.kariflag.value = "0";
  	// 2016/11/03 H.Yoshikawa Upd (IMDGChk追加)
  	retValue = showModalDialog("dmlModalAgree.asp?ChkAgr=" + chkagree + "&ChkSls=" + chksolas + "&BookChk=" + BookChk + "&IMDGChk=" + IMDGChk, window, "dialogWidth:550px; dialogHeight:250px; center:1; scroll: no; dialogTop:300px; ");
  	if(retValue != true){
  	  return false;
  	}
  }
//2016/08/02 H.Yoshikawa Add End

  return true;
}
<%'CW-017 ADD START%>
//ヘッドIDの制御
function checkIDF(type){
<% 'Change 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
   'If UpFlag <> 5 Then 
   If UpFlag <> 5 AND (Mord=0 OR Request("compFlag")) AND Request("TruckerFlag")<>1 Then%>
  target=document.dmi320F;
  targetCOMPcd=target.CMPcd<%=UpFlag%>;
  COMPcd="<%=Session.Contents("COMPcd")%>";
  checkID(type,target,targetCOMPcd,COMPcd);
<% End If %>
}
<%'CW-017 ADD END%>
//2008-01-31 Add-S M.Marquez
// 明細項目の活性・非活性
function fSetTabIndex(){
    var max = document.dmi320F.elements.length;
    var i ;
    for(i =0; i < max; i++){
        if (document.dmi320F.elements[i].type == "text" && document.dmi320F.elements[i].readOnly == true){
            document.dmi320F.elements[i].tabIndex = -1;
        }
    }
}
//2008-01-31 Add-E M.Marquez

//2008-01-31 Add-S G.Ariola
function finit(){
    document.dmi320F.CONsize.focus();
}
//2008-01-31 Add-E G.Ariola

//2016/08/02 H.Yoshikawa Add Start
//船名・次航の検索画面表示
function VslSelect(){
	winname="searchVsl";
	target=document.dmi320F;
	vslnm = target.ShipN.value;
  	retValue = window.open("dmlModalVslVoy.asp?tgt=dmi320F&VslNm="+vslnm+"&fldvn=ShipN&flddspvy=NextV&dspkbn=LD", winname, "width=600, height=600, menubar=no, toolbar=no, scrollbars=yes");
  	return true;
}

//ブッキング情報との比較
function BookingCheck(){
	var target = document.dmi320F;
	var ret = true;
	var elm;
	var ret2;
	
	if(SpaceDel(target.ShipN.value).toUpperCase() != SpaceDel("<%=ExcVslName %>").toUpperCase()){
		target.ShipN.style.backgroundColor = '#ffb6c1';
		target.ShipN.focus();
		ret = false;
	}else{
		target.ShipN.style.backgroundColor = '#ffffff';
		target.ShipN.value = "<%=ExcVslName %>";
	}
	if(SpaceDel(target.NextV.value).toUpperCase() != SpaceDel("<%=ExcVoyage %>").toUpperCase()){
		target.NextV.style.backgroundColor = '#ffb6c1';
		target.NextV.focus;
		ret = false;
	}else{
		target.NextV.style.backgroundColor = '#ffffff';
		target.NextV.value = "<%=ExcVoyage %>";
	}
	
	// 2016/11/04 H.Yoshikawa Add (if文追加)
	if("<%=ExcSize %>" != ""){
		if(target.CONsize.options[target.CONsize.selectedIndex].value != "<%=ExcSize %>"){
			target.CONsize.style.backgroundColor = '#ffb6c1';
			target.CONsize.focus();
			ret = false;
		}else{
			target.CONsize.style.backgroundColor = '#ffffff';
		}
		if(target.CONtype.options[target.CONtype.selectedIndex].value != "<%=ExcType %>"){
			target.CONtype.style.backgroundColor = '#ffb6c1';
			target.CONtype.focus();
			ret = false;
		}else{
			target.CONtype.style.backgroundColor = '#ffffff';
		}
		if(target.CONhite.options[target.CONhite.selectedIndex].value != "<%=ExcHeight %>"){
			target.CONhite.style.backgroundColor = '#ffb6c1';
			target.CONhite.focus();
			ret = false;
		}else{
			target.CONhite.style.backgroundColor = '#ffffff';
		}
	}else{
		target.CONsize.style.backgroundColor = '#ffffff';
		target.CONtype.style.backgroundColor = '#ffffff';
		target.CONhite.style.backgroundColor = '#ffffff';
	// 2016/11/04 H.Yoshikawa Add Start
	}
	
	retflg = false;
	<% for i = 0 to UBound(ExcSizeAry) %>
		if(target.CONsize.options[target.CONsize.selectedIndex].value == "<%=ExcSizeAry(i) %>"
		 && target.CONtype.options[target.CONtype.selectedIndex].value == "<%=ExcTypeAry(i) %>"
		 && target.CONhite.options[target.CONhite.selectedIndex].value == "<%=ExcHeightAry(i) %>")
		{
			retflg = true;
		}
	<% next %>
	if(retflg == false){
		target.CONsize.style.backgroundColor = '#ffb6c1';
		target.CONtype.style.backgroundColor = '#ffb6c1';
		target.CONhite.style.backgroundColor = '#ffb6c1';
		target.CONsize.focus();
		ret = false;
	}
	// 2016/11/04 H.Yoshikawa Add End

	// 2016/11/10 H.Yoshikawa Del Start
	//if(target.SttiT.readOnly == false){
	//	if(Number(Rtrim(target.SttiT.value, " ")) != Number("<%=ExcSetTemp %>")){
	//		target.SttiT.style.backgroundColor = '#ffb6c1';
	//		target.SttiT.focus();
	//		ret = false;
	//	}else{
	//		target.SttiT.style.backgroundColor = '#ffffff';
	//	}
	//}
	// 2016/11/10 H.Yoshikawa Del End
	
	// 2016/11/03 H.Yoshikawa Del Start
	//危険品関連チェック
	//for(k=1;k<=5;k++){
	//	ret2 = false;
	//	
	//	//危険品コード
	//	elm = target.elements["IMDG" + k];
	//	if(Rtrim(elm.value, " ") != ""){
	//		//ブッキング比較チェック
	//		if(Rtrim(elm.value, " ") != "<%=ExcIMDG1%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcIMDG2%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcIMDG3%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcIMDG4%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcIMDG5%>")
	//		{
	//			elm.style.backgroundColor = '#ffb6c1';
	//			elm.focus();
	//			ret = false;
	//		}else{
	//			//危険品コードチェック
	//			<% for i = 0 to UBound(mIMDG) %>
	//				if(Rtrim(elm.value, " ") == "<%=mIMDG(i)%>"){
	//					ret2 = true;
	//				}
	//			<% Next %>
	//			if(ret2 == false){
	//				elm.style.backgroundColor = '#ffb6c1';
	//				elm.focus();
	//				ret = false;
	//			}else{
	//				elm.style.backgroundColor = '#ffffff';
	//			}
	//		}
	//	}else{
	//		elm.style.backgroundColor = '#ffffff';
	//	}
	//	
	//	//UNNo
	//	elm = target.elements["UNNo" + k];
	//	if(Rtrim(elm.value, " ") != ""){
	//		//ブッキング比較チェック
	//		if(Rtrim(elm.value, " ") != "<%=ExcUNNo1%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcUNNo2%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcUNNo3%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcUNNo4%>"
	//		&& Rtrim(elm.value, " ") != "<%=ExcUNNo5%>")
	//		{
	//			elm.style.backgroundColor = '#ffb6c1';
	//			elm.focus();
	//			ret = false;
	//		}else{
	//			elm.style.backgroundColor = '#ffffff';
	//		}
	//		
	//	}else{
	//		elm.style.backgroundColor = '#ffffff';
	//	}
	//}
	// 2016/11/03 H.Yoshikawa Del End
	
	return ret;
}

//2016/08/02 H.Yoshikawa Add End

// 2016/11/03 H.Yoshikawa Add Start
//IMDGコードのマスタとの比較
function IMDGCheck(){
	var target = document.dmi320F;
	var ret = true;
	var elm;
	var ret2;

	//危険品関連チェック
	for(k=1;k<=5;k++){
		ret2 = false;
		
		//危険品コード
		elm = target.elements["IMDG" + k];
		if(Rtrim(elm.value, " ") != ""){
			//危険品コードチェック
			<% for i = 0 to UBound(mIMDG) %>
				if(Rtrim(elm.value, " ") == "<%=mIMDG(i)%>"){
					ret2 = true;
				}
			<% Next %>
			if(ret2 == false){
				elm.style.backgroundColor = '#ffb6c1';
				elm.focus();
				ret = false;
			}else{
				elm.style.backgroundColor = '#ffffff';
			}
		}else{
			elm.style.backgroundColor = '#ffffff';
		}
	}
	
	return ret;
}

//港の検索画面表示
function PortSelect(portsbt){
	//winname="searchPort";
	target=document.dmi320F;
	
	if(portsbt == "Niuke"){
		codefld = "NiukP";
		namefld = "NiukeNm";
	}else if(portsbt == "LPort"){
		codefld = "TumiP";
		namefld = "LPortNm";
	}else if(portsbt == "DPort"){
		codefld = "AgeP";
		namefld = "DPortNm";
	}else if(portsbt == "Niwata"){
		codefld = "NiwataP";
		namefld = "NiwataNm";
	}
	
  	retValue = window.open("dmlModalPort.asp?tgt=dmi320F&fldcode="+codefld+"&fldname="+namefld, "", "width=400, height=600, menubar=no, toolbar=no, scrollbars=yes");
  	return true;
}
// 2016/11/03 H.Yoshikawa Add End

// 2017/04/04 H.Yoshikawa Add Start
//確定事業者の検索画面表示
function DfTSelect(){
	target=document.dmi320F;
	dftcd = target.ReportNo.value;
  	retValue = window.showModalDialog("dmlModalDefTrade.asp?DfTCd="+dftcd, window, "dialogWidth:0px; dialogHeight:0px; center:1; scroll: no;");
  	
  	target.DefName.value = retValue;

  	if(target.ReportNo.value.length != 0 && target.DefName.value.length == 0){
	    alert("入力された届出番号は、確定事業者マスタに登録されていません。");
	    target.ReportNo.focus();
	    return false;
  	}
  	return true;
}
// 2017/04/04 H.Yoshikawa Add End

// -->

function CheckKana(str){
  checkstr="｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ";
   for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) >= 0){
      return false;
    }
  }
  return true;
}
//2009/07/27 Add-S C.Pestano
function CheckLen(obj,mesgon,focuson,mandatory) {
	var kanjicheck = gfStrLen(obj.value);
	
	if (kanjicheck == false){
		alert("半角文字を入力してください。");
		obj.focus();
		return false;
	}	
	
	if (mandatory && objlength==0)
		return false;	
	return true;
}

function gfStrLen(StrSrc) {
	var r = 0;
	for (var i = 0; i < StrSrc.length; i++) {
		var c = StrSrc.charCodeAt(i);
		// Shift_JIS: 0x0 ～ 0x80, 0xa0  , 0xa1   ～ 0xdf  , 0xfd   ～ 0xff
		// Unicode  : 0x0 ～ 0x80, 0xf8f0, 0xff61 ～ 0xff9f, 0xf8f1 ～ 0xf8f3
		if ( (c >= 0x0 && c < 0x81) || (c == 0xf8f0) || (c >= 0xff61 && c < 0xffa0) || (c >= 0xf8f1 && c < 0xf8f4)) {
			
		} else {			
			return false;		
		}
	}
	return true;
}
//2009/07/27 Add-E C.Pestano

//2017/03/02  T.Okui Add-S
//文字列に数字が含まれているかチェック
function numCheck(str){
	var flg = 0;
	var num1 = ['0','1','2','3','4','5','6','7','8','9'];
	var num2 = ['０','１','２','３','４','５','６','７','８','９'];
	for(var i=0;i < 10 ;i++){
		if(str.indexOf(num1[i]) >= 0){
			flg = 1;
			break;
		}
		if(str.indexOf(num2[i]) >= 0){
			flg = 1;
			break;
		}
	}
	if(flg == 1){
		return true;
	}
	return false;
}


//2017/03/02  T.Okui Add-E
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<!-- 2016/08/08 H.Yoshikawa Upd Start -->
<!-- <BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi320F);fSetTabIndex();finit();"> -->
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmi320F);fSetTabIndex();finit();">
<!-- 2016/08/08 H.Yoshikawa Upd End -->
<!-------------実搬入情報入力(更新)画面--------------------------->
<FORM name="dmi320F" method="POST">

<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <% If Mord<>"0" Then %>
	    <TD><B>搬入票作成情報入力(更新モード)</B></TD>
	<% Else %>
	    <TD><B>搬入票作成情報入力</B></TD>
	<% End If %>
    <TD colspan=2>
    </TD></TR>
  <TR>
    <TD width="500" colspan=2 valign=top>
    <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	  <TD>
	<DIV style="height:330px;width:500px;border: 1px solid black; margin:5px;">
	<TABLE border=0 cellPadding=2 cellSpacing=0>
	  <TR>
    	    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">BOOKING情報</DIV></TD>
    	    <TD></TD>
      </TR>
	  <TR>
	    <TD><DIV class=bgb>ブッキング番号</DIV></TD>
	    <TD><INPUT type=text name="BookNo" value="<%=Request("BookNo")%>" readOnly></TD>
	  </TR>
	  <TR>
	    <TD width="90px"><DIV class=bgb>コンテナ番号</DIV></TD>
	    <TD><INPUT type=text name="CONnum" value="<%=CONnum%>" readOnly tabindex=-1></TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>*シール番号</DIV></TD>
	    <TD><INPUT type=text name="SealNo" value="<%=Request("SealNo")%>" maxlength=15></TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>取扱船社</DIV></TD>
	    <TD><INPUT type=text name="ShipLineName" value="<%=Request("ShipLineName")%>" readOnly size=40>
	    	<INPUT type=hidden name="ThkSya" value="<%=Request("ThkSya")%>">
	    </TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>*本船名</DIV></TD>
	    <TD><INPUT type=text name="ShipN" value="<%=Request("ShipN")%>" maxlength=20>	<!-- 2016/08/01 H.Yoshikawa Upd （readOnly属性削除) -->
	    	<INPUT type=button value="検索" onClick="VslSelect()">
	    	<INPUT type=hidden name="ShipCode" value="<%=Request("ShipCode")%>">		<!-- 2016/08/01 H.Yoshikawa Add -->
	    </TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>*次航</DIV></TD>
	    <TD><INPUT type=text name="NextV" value="<%=Request("NextV")%>" maxlength=12></TD>			<!-- 2016/08/01 H.Yoshikawa Upd （readOnly属性削除) -->
	    <TD><INPUT type=hidden name="VoyCtrl" value="<%=Request("VoyCtrl")%>" maxlength=12></TD>	<!-- 2016/10/14 H.Yoshikawa Add -->
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>*荷受地</DIV></TD>
	    <TD><INPUT type=text name="NiukP" value="<%=Request("NiukP")%>" maxlength=5 size=8 placeholder="港コード">				<!-- 2016/11/02 H.Yoshikawa Upd （readOnly属性削除、maxlength,size,placeholder追加) -->
	    	<INPUT type=text name="NiukeNm" value="<%=Request("NiukeNm")%>" size=30 readOnly  placeholder="名称">				<!-- 2016/11/03 H.Yoshikawa Add -->
			<INPUT type=button value="検索" onClick="PortSelect('Niuke')">
	    </TD>
	  </TR>
	  
	  <TR>
	    <TD><DIV class=bgb>*積港</DIV></TD>
	    <TD><INPUT type=text name="TumiP" value="<%=Request("TumiP")%>" maxlength=5 size=8 placeholder="港コード">				<!-- 2016/11/02 H.Yoshikawa Upd （readOnly属性削除、maxlength,size,placeholder追加) -->
	    	<INPUT type=text name="LPortNm" value="<%=Request("LPortNm")%>" size=30 readOnly  placeholder="名称">				<!-- 2016/11/03 H.Yoshikawa Add -->
			<INPUT type=button value="検索" onClick="PortSelect('LPort')">
	    </TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>*揚港</DIV></TD>
	    <TD><INPUT type=text name="AgeP" value="<%=Request("AgeP")%>" maxlength=5 size=8 placeholder="港コード">				<!-- 2016/11/02 H.Yoshikawa Upd （readOnly属性削除、maxlength,size,placeholder追加) -->
	    	<INPUT type=text name="DPortNm" value="<%=Request("DPortNm")%>" size=30 readOnly  placeholder="名称">				<!-- 2016/11/03 H.Yoshikawa Add -->
			<INPUT type=button value="検索" onClick="PortSelect('DPort')">
	    </TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>*荷渡地</DIV></TD>
	    <TD><INPUT type=text name="NiwataP" value="<%=Request("NiwataP")%>" maxlength=5 size=8 placeholder="港コード">			<!-- 2016/11/02 H.Yoshikawa Upd （readOnly属性削除、maxlength,size,placeholder追加) -->
	    	<INPUT type=text name="NiwataNm" value="<%=Request("NiwataNm")%>" size=30 readOnly  placeholder="名称">				<!-- 2016/11/03 H.Yoshikawa Add -->
			<INPUT type=button value="検索" onClick="PortSelect('Niwata')">
	    </TD>
	  </TR>
	<!-- 2016/08/01 H.Yoshikawa Add Start -->
	  <TR>
	    <TD><DIV class=bgb>*荷主</DIV></TD>
	    <TD><INPUT type=text name="Shipper" value="<%=Request("Shipper")%>" maxlength=80 size=40></TD>
	  </TR>
	  
	  <TR>
	    <TD><DIV class=bgb>搬入先</DIV></TD>
	    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>" readOnly size=30></TD>
	  </TR> 
	  <TR>
	    <TD><DIV class=bgb>ターミナルオペレータ</DIV></TD>
	    <!-- 2017/03/02 T.Okui Upd-S -->
    	<!--
	    <TD><INPUT type=text name="Operator" value="<%=Request("Operator")%>" readOnly></TD></TR>
	    -->
	    <TD><INPUT type=text name="OpeName" value="<%=gfConvertOperator(Request("Operator"))%>" readOnly></TD>
	    <INPUT type=hidden name="Operator" value="<%=Request("Operator")%>">
	    </TR>
	    <!-- 2017/03/02 T.Okui Upd-E -->
	  </TABLE>
	  </DIV>
	  </TD>
	  </TR>
	  <TR>
	  <TD>
	  <DIV style="height:160px;width:500px;border: 1px solid black; margin:5px;">
	  <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">登録情報</DIV></TD>
	    <TD></TD>
	  <TR>
	  <TR>
	    <TD><DIV class=bgb>登録会社コード</DIV></TD>
	    <TD>
	        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7 readOnly>
	    </TD>
	  </TR>
	  <!-- 2017/04/04 H.Yoshikawa Add Start -->
	  <TR>
	  	<TD><DIV class=bgb>指示元担当者</DIV></TD>
		<TD>
	    	<INPUT type=text name="EntryName" value="<%=Request("EntryName")%>" maxlength=20>
	    </TD>
	  </TR>
	  <!-- 2017/04/04 H.Yoshikawa Add End -->
	  <TR>
	    <TD><DIV class=bgb>指示先会社コード</DIV></TD>
	    <TD>
	        <INPUT type=text name="CMPcd1" value=<%=CMPcd(1)%> size=5 maxlength=2>
	        <INPUT type=text name="CMPcd2" value=<%=CMPcd(2)%> size=5 maxlength=2>
	        <INPUT type=text name="CMPcd3" value=<%=CMPcd(3)%> size=5 maxlength=2>
	        <INPUT type=text name="CMPcd4" value=<%=CMPcd(4)%> size=5 maxlength=2>
	    </TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>ヘッドＩＤ</DIV></TD>
	<!-- CW-017 Chenge
	    <TD><INPUT type=text name="HedId" value="<%=HedId%>")"></TD>
	-->
	    <TD><INPUT type=text name="HedId" value="<%=HedId%>" maxlength=5 onBlur="checkIDF(1)"></TD>
	  </TR>
	  <TR>
		<TD><DIV class=bgb>*搬入元</DIV></TD>
		<TD><INPUT type=text name="HFrom" value="<%=Request("Hfrom")%>" size=35 maxlength=30></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>搬入予定日</DIV></TD>
	    <TD><INPUT type=text name="Hmon" value="<%=Request("Hmon")%>" size=2 maxlength=2>月
	        <INPUT type=text name="Hday" value="<%=Request("Hday")%>" size=2 maxlength=2>日</TD>
	  </TR>
	  </TABLE>
	  </DIV>
	  </TD>
	  </TR>
	  
	  <TR><TD colspan=2>
	  	<DIV style="height:100px;width:500px;border: 1px solid black; margin:5px;">
	  	<table border=0 cellPadding=2 cellSpacing=0>
	  	<TR>
	    <TD colspan=2><DIV><SPAN  STYLE="FONT-WEIGHT:BOLD;">搬入票署名欄情報</SPAN><SPAN STYLE="color:red;">※この欄が搬入票署名欄に印字されます</SPAN></DIV></TD>
	    </TR>
	  	<TR><TD><DIV class=bgb>*取扱海貨社名<BR></DIV></TD>
			<TD><INPUT type=text name="Forwarder" value="<%=Request("Forwarder")%>" style="margin-bottom:2px;" maxlength=80 size=40>
	    	</TD>
		</TR>
		<TR><TD><DIV class=bgb>*（担当者）</DIV></TD>
			<TD>
	    		<INPUT type=text name="FwdStaff" value="<%=Request("FwdStaff")%>" maxlength=20>
	    	</TD>
		</TR>
		<TR><TD><DIV class=bgb>*（連絡先）</DIV></TD>
			<TD>
	    		<INPUT type=text name="FwdTel" value="<%=Request("FwdTel")%>" maxlength=15>
	    	</TD>
		</TR>
		</table>
		</div>
	  </TD></TR>
	<!-- 2016/08/01 H.Yoshikawa Add End   -->
	<TR>
	  <TD>
	  <DIV style="height:55px;width:500px;border: 1px solid black; margin:5px;">
	  <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  	<TR>
	   <TD><DIV class=bgy>*登録担当者</DIV></TD>
	   <TD><INPUT type=text name="TruckerSubName" value="<%=TruckerSubName%>"  maxlength=16 autocomplete="on"></TD>
	<!-- 2009/03/10 R.Shibuta Add-E -->
	  </TR>
<!-- 2016/08/18 H.Yoshikawa Add-S -->
	  <TR>
	   <TD><DIV class=bgy>*登録者連絡先</DIV></TD>
	   <TD><INPUT type=text name="TruckerTel" value="<%=Request("TruckerTel")%>"  maxlength=15 onBlur="CheckLen(this,true,true,false)" autocomplete="on"></TD>
	  </TR>
	<!-- 2016/08/17 H.Yoshikawa Add End   -->
	</TABLE>
	</DIV>
	</TD>
	</TR>
	</TABLE>
</TD>
<TD width=300 valign=top>
	<TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  
	  <TR>
	    <TD width=300>
	    <DIV style="height:150px;width:300px;border: 1px solid black; margin:5px;">
		    <TABLE cellpadding=1 cellspacing=0>
		    <TR>
		    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">コンテナ情報</DIV></TD>
		    <TD></TD></TR>
		    <TR>
		    <TD width=126>
	        	<DIV class=bgb>*サイズ</DIV></TD>
		    <TD><select name="CONsize" style="width:47px"></select>			<!-- 2017/04/20 H.Yoshikawa Upd(style="width:47px"追加) -->
		    </TD></TR>
		    <TR>
		    <TD width=126>
	        <DIV class=bgb>*タイプ</DIV></TD>
		    <TD>
		        <select name="CONtype" style="width:47px"></select>			<!-- 2017/04/20 H.Yoshikawa Upd(style="width:47px"追加) -->
		    </TD></TR>
	    <TR>
		    <TD width=126>
	        <DIV class=bgb>*高さ</DIV></TD>
	    <TD>
	        <select name="CONhite" style="width:47px"></select>				<!-- 2017/04/20 H.Yoshikawa Upd(style="width:47px"追加) -->
	    </TD></TR>
	    <!-- 20170118 T.OKui Upd Start -->
	   <TR>
	    <TD><DIV class=bgb>設定温度</DIV></TD>
	    <TD><INPUT type=text name="SttiT" value="<%=ExcSetTemp%>" size=6 maxlength=5 readOnly>&nbsp;<%=TempDegree%>
	    <INPUT type=hidden name="TempDegree" value="<%=TempDegree %>">
	    <!--
	    	<INPUT type=checkbox name="AsDry" value="1" <% if gfTrim(Request("AsDry")) = "1" then %>checked<% end if %>>AS DRY
	    -->
	    </TD>
	  </TR>
	  <!-- 20170118 T.OKui Upd Start -->
	  <TR>
	    <TD><DIV class=bgb><!--2017/04/04 H.Yoshikawa Upd ＶＥＮＴ -->ベンチレーション</DIV></TD>
	    <TD><INPUT type=text name="VENT" value="<%=Request("VENT")%>" size=5 maxlength=3>&nbsp;%（開口）</TD></TR>	<!-- 2016/08/01 H.Yoshikawa Upd （readOnly属性削除) -->
	  <!-- 2017/03/02 T.Okui Del Start -->
	  <!--
	  <TR>
	    <TD><DIV class=bgb>丸関</DIV></TD>  -->
	    <!-- 2016/11/16 H.Yoshikawa Upd Start -->
	    <!-- <TD><select name="MrSk">
	          <OPTION value=" "> 
	          <OPTION value="Y">Y
	          <OPTION value="N">N
	        </select></TD></TR> -->
	    <!--
	    <TD><INPUT type=text name="MrSk" value="<%=Request("MrSk")%>" size=5 ReadOnly></TD></TR>-->
	    <!-- 2016/11/16 H.Yoshikawa Upd End -->
	    <!-- 2017/03/02 T.Okui Del End -->
	  </TABLE>
		    </DIV>
	    </TD>
	    <TD valign=top>
	     	<DIV style="height:150px;width:300px;border: 1px solid black; margin:5px;">
		  	<TABLE cellpadding=1 cellspacing=0>
		  		<TR>
			    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">コンテナ規格外貨物情報</DIV></TD>
			    <TD></TD></TR>    
			    <TR>
			    <TD><DIV class=bgb>オーバーハイ（上部）</DIV></TD>
			    <TD><INPUT type=text name="OH"  value="<%=Request("OH")%>"  size=5 maxlength=7>&nbsp;cm</TD></TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーワイド（右）</DIV></TD>
			    <TD><INPUT type=text name="OWR" value="<%=Request("OWR")%>" size=5 maxlength=7>&nbsp;cm</TD></TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーワイド（左）</DIV></TD>
			    <TD><INPUT type=text name="OWL" value="<%=Request("OWL")%>" size=5 maxlength=7>&nbsp;cm</TD></TR>
			    <TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーレンクス（前）</DIV></TD>													<!-- 2017/04/20 H.Yoshikawa Upd(レングス→レンクス) --> 
			    <TD><INPUT type=text name="OLF" value="<%=Request("OLF")%>" size=5 maxlength=7>&nbsp;cm</TD></TR>
			    <TR>
			    <TD><DIV class=bgb>オーバーレンクス（後）</DIV></TD>													<!-- 2017/04/20 H.Yoshikawa Upd(レングス→レンクス) --> 
			    <TD><INPUT type=text name="OLA" value="<%=Request("OLA")%>" size=5 maxlength=7>&nbsp;cm</TD></TR>
			    </TR>
			</TABLE></DIV>
	    </TD>
	  </TR>  
	  <TR>
	  <TD colspan=2>
	  <DIV style="border: 1px solid black; margin:5px;height:145px;">
	  <TABLE>
	  <TR>
	    <TD width="115"><DIV STYLE="FONT-WEIGHT:BOLD;">重量情報</DIV></TD>
	    <TD></TD></TR>
	  <TR>
	    <TD style="padding-bottom:0px;"><DIV class=bgb>*コンテナ総重量</DIV></TD>
	    <TD style="padding-bottom:0px;"><INPUT type=text name="GrosW" value="<%=Request("GrosW")%>" size=9 maxlength=8>&nbsp;kg</TD>
	  </TR>
	  <TR>
	    <TD width=126>
	        <DIV class=bgb>テアウェイト</DIV></TD>
	    <TD><!-- 2016/08/01 H.Yoshikawa Delete 
	        <select name="CONsitu"></select> -->
	        <!--  2017/03/02 T.Okui Del
	        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=7 ReadOnly>kg-->			<!-- 2016/11/17 H.Yoshikawa Upd(readonly) --> 
	        <!--  2017/03/02 T.Okui Upd-S-->
	        <!-- 2017/03/02 T.Okui 新規登録時はKACCSの値を表示。readonlyを外す -->
	        <% if Mord="0" then%>
	        <INPUT type=text name="CONtear" value="<%=TareWeight%>" maxlength=4 size=9>&nbsp;kg					<!-- 2017/04/20 H.Yoshikawa Upd(size=7→9、kgの前にスペース追加) --> 
	        <% else %>
	        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" maxlength=4 size=9>&nbsp;kg			<!-- 2017/04/20 H.Yoshikawa Upd(size=7→9、kgの前にスペース追加) --> 
	        <% end if %>
	        <!--  2017/03/02 T.Okui Upd-E-->
	    </TD>
	  </TR>
	  <TR>
	  	<TD><DIV class=bgb>計量方法（確認）</DIV></TD>
		<TD style="padding-top:0px;"><INPUT type=checkbox name="SolasChk" value="1" <% if gfTrim(Request("SolasChk")) = "1" then %>checked<% end if %>>ここに入力したコンテナ総重量はSOLAS条約に基づく方法で計測された数値です。</TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>届出番号　登録番号</DIV></TD>
	    <!--<TD><INPUT type=text name="ReportNo" value="<%=Request("ReportNo")%>" size=25 maxlength=20 title="荷送人の届出番号、または登録確定事業者番号を入力してください。">　<span style="color:red">※ハイフンなしで入力してください</span></TD></TR>-->
	    <TD><INPUT type=text name="ReportNo" value="<%=Request("ReportNo")%>" size=17 maxlength=12 title="荷送人の届出番号、または登録確定事業者番号を入力してください。">
	        <INPUT type=button value="検索" onClick="DfTSelect();">		<!-- 2017/04/04 H.Yoshikawa Add -->
	        　<span style="color:red">※ハイフンなしで入力してください</span></TD></TR>
	  <!-- 2017/04/04 H.Yoshikawa Add Start -->
	  <TR>
	    <TD><DIV class=bgb>重量　確定事業者</DIV></TD>
	    <TD><INPUT type=text name="DefName" value="<%=Request("DefName")%>" size=40 readOnly></TD></TR>
	  <!-- 2017/04/04 H.Yoshikawa Add End -->
	  <!-- 2016/08/02 H.Yoshikawa Delete  
	  <TR>
	    <TD><DIV class=bgb>*通関</DIV></TD>
	    <TD><select name="TuSk">
	          <OPTION value=" "> 
	          <OPTION value="済">済
	          <OPTION value="未">未
	        </select></TD></TR>	-->
	  </TABLE>
	  </DIV>
	  </TD>
	  </TR>
	  <TR>
	  <TD colspan=2>
	  <DIV style="border: 1px solid black; margin:5px;height:205px;">
	  <TABLE>
	  <TR>
	    <TD><DIV STYLE="FONT-WEIGHT:BOLD;">危険品貨物情報</DIV></TD>
	    <TD></TD></TR>
	  <TR>
	    <!-- 2017/04/04 H.Yoshikawa Del <TD><DIV class=bgb>ＲＨＯ</DIV></TD> -->
	    <TD><INPUT type=hidden name="RHO" value="<%=Request("RHO")%>" size=5 readOnly></TD></TR> <!-- 2017/04/04 H.Yoshikawa Upd text→hidden -->
	  <TR>
		    <TD></TD>
		    <TD>
	        	<TABLE cellpadding=0 cellspacing=0 style="font-weight:bold;"><TR>
		        	<TD align=center width="47px">1</TD>
		        	<TD align=center width="50px">2</TD>
		        	<TD align=center width="50px">3</TD>
		        	<TD align=center width="50px">4</TD>
		        	<TD align=center width="50px">5</TD>
		        </TR></TABLE>
		    </TD></TR>
	  
	  <TR>
	    <TD><DIV class=bgb>ＩＭＤＧコード</DIV></TD>
	    <TD>
	        <TABLE cellpadding=0 cellspacing=0><TR>
	        	<TD width="50px"><INPUT type=text name="IMDG1" value="<%=Request("IMDG1")%>" size=6 maxlength=3></TD>	<!-- 2016/08/01 H.Yoshikawa Upd （size5→6、readOnly属性削除) -->
	        	<TD width="50px"><INPUT type=text name="IMDG2" value="<%=Request("IMDG2")%>" size=6 maxlength=3></TD>	<!-- 2016/08/01 H.Yoshikawa Upd （size5→6、readOnly属性削除) -->
	        	<TD width="50px"><INPUT type=text name="IMDG3" value="<%=Request("IMDG3")%>" size=6 maxlength=3></TD>	<!-- 2016/08/01 H.Yoshikawa Upd （size5→6、readOnly属性削除) -->
	        	<TD width="50px"><INPUT type=text name="IMDG4" value="<%=Request("IMDG4")%>" size=6 maxlength=3></TD>	<!-- 2016/08/01 H.Yoshikawa Add -->
	        	<TD width="50px"><INPUT type=text name="IMDG5" value="<%=Request("IMDG5")%>" size=6 maxlength=3></TD>	<!-- 2016/08/01 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>ＵＮコード</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="UNNo1" value="<%=Request("UNNo1")%>" size=6 maxlength=4></TD>					<!-- 2016/08/01 H.Yoshikawa Upd （readOnly属性削除) -->
		        <TD width="50px"><INPUT type=text name="UNNo2" value="<%=Request("UNNo2")%>" size=6 maxlength=4></TD>					<!-- 2016/08/01 H.Yoshikawa Upd （readOnly属性削除) -->
		        <TD width="50px"><INPUT type=text name="UNNo3" value="<%=Request("UNNo3")%>" size=6 maxlength=4></TD>					<!-- 2016/08/01 H.Yoshikawa Upd （readOnly属性削除) -->
		        <TD width="50px"><INPUT type=text name="UNNo4" value="<%=Request("UNNo4")%>" size=6 maxlength=4></TD>					<!-- 2016/08/01 H.Yoshikawa Add -->
		        <TD width="50px"><INPUT type=text name="UNNo5" value="<%=Request("UNNo5")%>" size=6 maxlength=4></TD>					<!-- 2016/08/01 H.Yoshikawa Add -->
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <!-- 2016/08/01 H.Yoshikawa Add Start -->
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>サブラベル１</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="Label1" value="<%=Request("Label1")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="Label2" value="<%=Request("Label2")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="Label3" value="<%=Request("Label3")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="Label4" value="<%=Request("Label4")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="Label5" value="<%=Request("Label5")%>" size=6 maxlength=3></TD>
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>サブラベル２</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px"><INPUT type=text name="SubLabel1" value="<%=Request("SubLabel1")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel2" value="<%=Request("SubLabel2")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel3" value="<%=Request("SubLabel3")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel4" value="<%=Request("SubLabel4")%>" size=6 maxlength=3></TD>
		        <TD width="50px"><INPUT type=text name="SubLabel5" value="<%=Request("SubLabel5")%>" size=6 maxlength=3></TD>
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <!-- 2016/08/01 H.Yoshikawa Add End -->
	  
	  <!-- 2016/08/01 H.Yoshikawa Add Start -->
	  <TR>
	    <TD style="padding-top:0px;"><DIV class=bgb>少量危険品</DIV></TD>
	    <TD style="padding-top:0px;">
	        <TABLE cellpadding=0 cellspacing=0><TR>
		        <TD width="50px" align=center><INPUT type=checkbox name="LqFlag1" value="1" <% if gfTrim(Request("LqFlag1")) = "1" then %>checked<% end if %>></TD>
		        <TD width="50px" align=center><INPUT type=checkbox name="LqFlag2" value="1" <% if gfTrim(Request("LqFlag2")) = "1" then %>checked<% end if %>></TD>
		        <TD width="50px" align=center><INPUT type=checkbox name="LqFlag3" value="1" <% if gfTrim(Request("LqFlag3")) = "1" then %>checked<% end if %>></TD>
		        <TD width="50px" align=center><INPUT type=checkbox name="LqFlag4" value="1" <% if gfTrim(Request("LqFlag4")) = "1" then %>checked<% end if %>></TD>
		        <TD width="50px" align=center><INPUT type=checkbox name="LqFlag5" value="1" <% if gfTrim(Request("LqFlag5")) = "1" then %>checked<% end if %>></TD>
	        </TR></TABLE>
	    </TD>
	  </TR>
	  <!-- 2016/08/01 H.Yoshikawa Add End -->
	  </TABLE>
	  </TD></TR>
	  
	  
	  <TR>
	 
	<TD colspan=2 valign="TOP">
	<TABLE border=0 cellpadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD align="left" valign="top">
	    　＜注意事項＞<BR>
	    　本画面の誤記・記入漏れは正常なる輸送を阻害しますので、入力済みの項目も含めて必ずご確認ください。<BR>
        　誤記・記入漏れにより発生する損害・費用・罰金等は全て本画面の入力者が負担し、船社(含むターミナル)<BR>
        　は責任を負いません。
	    </TD>
	  </TR>
	</TABLE>
	</TD>
  </TR>
  <TR>
  <TD colspan=2 valign="TOP">
  <TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
	  <TR>
	    <TD colspan=3 align=left valign=bottom>
	        <BR><INPUT type=checkbox name="AgreeChk" value="1" <% if gfTrim(Request("AgreeChk")) = "1" then %>checked<% end if %>><!-- 2017/04/04 H.Yohsikawa Upd 文字色：赤 --><span style="color:red">本画面の入力内容をゲートでの搬入票の代わりとして使用することに同意します。</span>
	    	<BR>　<!-- 2017/04/04 H.Yohsikawa Upd 文字色：赤 --><span style="color:red">※チェックを入れずに「登録」をした場合は、仮登録であり、予約受付は完了していません。</span>
	    	
	    	
		</TD>
	  </TR>
	<!-- 2016/08/18 H.Yoshikawa Add-E -->
	</TABLE>
	</TD>
	</TR>  
	<!-- 2016/08/01 H.Yoshikawa Add-E -->
	<TR>
	 <TD colspan=2 align=center><BR/><BR/>
	   <INPUT type=hidden name="SakuNo"   value="<%=Request("SakuNo")%>">
	   <INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>">
	   <INPUT type=hidden name="UpFlag"   value="<%=UpFlag%>">
	   <INPUT type=hidden name="Mord"     value="<%=Mord%>" >
	   <INPUT type=hidden name="partFlg"  value="<%=partFlg%>" >
	   <INPUT type=hidden name="TruckerFlag"  value="<%=Request("TruckerFlag")%>" >
	   <INPUT type=hidden name="kariflag" value="">					<!-- 2016/10/12 H.Yoshikawa Add -->
<!-- 2016/11/03 H.Yoshikawa Upd Start -->

<% If ErrMsg <> "" Then %>
       <DIV class=alert>
        <%=ErrMsg%>
       </DIV>
       <BR>
	   <INPUT type=button value="削除" onClick="GoDell()" <% 'style="position:relative;left:50px;" %>>　
	   <INPUT type=button value="キャンセル" onClick="window.close()" <% 'style="position:relative;left:80px;" %>>
<% Else %>
	<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
	  ' If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then
	   If Request("TruckerFlag")<>1 AND UpFlag <> 1 AND Request("compFlag") Then %>
	       <DIV class=bgw>指示元へ回答　　　
	       <INPUT type=radio name="way" checked>Yes　
	       <INPUT type=radio name="way">No</DIV>
	<% End If %>
	<% If Mord="0" Then %>
	       <INPUT type=button value="登録" onClick="GoReEntry()">
	<% Else %>
	  <%'20030909 IF Request("TruckerFlag")<>1 Then %>
	<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
	  '   IF Request("TruckerFlag")<>1 AND Request("compFlag") Then 
	     IF Request("compFlag") Then %>
	       <INPUT type=button value="更新" onClick="GoReEntry()">&nbsp;
	  <% End If %>
	  <% IF UCase(Session.Contents("userid"))=CMPcd(0) Then %>
	       <INPUT type=hidden name=WkCNo value="<%=Request("WkCNo")%>" >
	       <INPUT type=button value="削除" onClick="GoDell()">&nbsp;
	  <% End If %>
	<% End If %>
		   <INPUT type=button value="コンテナ情報" onClick="GoConInfo()">&nbsp;
	       <INPUT type=button value="キャンセル" onClick="window.close()">
<% End If %>
<!--2017/02/06 T.Okui Del End-->
	    <% If Mord<>"0" Then %>
	      <TABLE border=1 cellPadding=3 cellSpacing=0 align="left" <% '2017/04/04 H.Yoshikawa Upd style="position:relative;left:6px;" %> style="float:right;">
	          <TR bgcolor="#f0f0f0"><TD>作業番号</TD><TD><%=SakuNo%></TD></TR>
	      </TABLE>
	      <!--<span style="padding:3px;border: 1px solid black; background-color:#f0f0f0;">作業番号|<%=SakuNo%></span> -->
	    <% End If %>
	  </TD>
	</TR>
  </TABLE>
  </TD>
  </TR>
  
  <!--2017/02/06 T.Okui Del Start-->
  <% if 1=0 then%>
	</TABLE>
</TD>
</TR>
<TR>
<TD valign="TOP">
	<TABLE border=0 cellPadding=2 cellSpacing=0>
	  <TR>
	<!-- 2009/03/10 R.Shibuta Add-S -->
	   <TD><DIV class=bgy>*登録担当者</DIV></TD>
	   <!-- 2009/07/25 Update C.Pestano -->
	   <TD><INPUT type=text name="TruckerSubName" value="<%=TruckerSubName%>"  maxlength=16 autocomplete="on"></TD>
	<!-- 2009/03/10 R.Shibuta Add-E -->
	  </TR>
	<!-- 2016/08/01 H.Yoshikawa Add-S -->
	  <TR>
	   <TD><DIV class=bgy>*登録者連絡先</DIV></TD>
	   <TD><INPUT type=text name="TruckerTel" value="<%=Request("TruckerTel")%>"  maxlength=15 onBlur="CheckLen(this,true,true,false)" autocomplete="on"></TD>
	  </TR>
	  <!--
	  <TR>
	    <TD width="90px"><DIV class=bgb>備考１</DIV></TD>
	    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73></TD>
	  </TR>
	  <TR>
	    <TD><DIV class=bgb>備考２</DIV></TD>
	    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73></TD></TR>
	  <TR>
	    <TD><DIV class=bgb>備考３</DIV></TD>
	    <TD><INPUT type=text name="Comment3" value="<%=Request("Comment3")%>" size=73></TD></TR>-->
	</TABLE>
</TD>
<TD>
	<TABLE border=0 cellPadding=2 cellSpacing=0>
	  <TR>
	<!-- 2016/08/01 H.Yoshikawa Add-S -->
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
	<!-- 2016/08/01 H.Yoshikawa Add-E -->
</TD>
</TR>
<TR>
<TD colspan=2 align=center>
<BR>
	<TABLE border=0 cellPadding=2 cellSpacing=0>
	  <TR>
	    <TD colspan=3 align=center valign=bottom>
	    	<BR><INPUT type=checkbox name="AgreeChk" value="1" <% if gfTrim(Request("AgreeChk")) = "1" then %>checked<% end if %>>本画面の入力内容をゲートでの搬入票の代わりとして使用することに同意します。
	    	<BR>　※チェックを入れずに「登録」をした場合は、仮登録であり、予約受付は完了していません。
		</TD>
	  </TR>
	<!-- 2016/08/01 H.Yoshikawa Add-E -->
	  <TR>
	    <TD colspan=3 align=center valign=bottom>
	       <INPUT type=hidden name="SakuNo"   value="<%=Request("SakuNo")%>">
	       <INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>">
	       <INPUT type=hidden name="UpFlag"   value="<%=UpFlag%>">
	       <INPUT type=hidden name="Mord"     value="<%=Mord%>" >
	       <INPUT type=hidden name="partFlg"  value="<%=partFlg%>" >
	       <INPUT type=hidden name="TruckerFlag"  value="<%=Request("TruckerFlag")%>" >
	       <INPUT type=hidden name="kariflag" value="">					<!-- 2016/10/12 H.Yoshikawa Add -->
<!-- 2016/11/03 H.Yoshikawa Upd Start -->
<% If ErrMsg <> "" Then %>
       <DIV class=alert>
        <%=ErrMsg%>
       </DIV>
	   <INPUT type=button value="削除" onClick="GoDell()">　　
	   <INPUT type=button value="キャンセル" onClick="window.close()">
<% Else %>
	<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
	  ' If Request("TruckerFlag")<>1 AND UpFlag <> 1 Then
	   If Request("TruckerFlag")<>1 AND UpFlag <> 1 AND Request("compFlag") Then %>
	       <DIV class=bgw>指示元へ回答　　　
	       <INPUT type=radio name="way" checked>Yes　
	       <INPUT type=radio name="way">No</DIV><P>
	<% End If %>
	<% If Mord="0" Then %>
	       <INPUT type=button value="登録" onClick="GoReEntry()">
	<% Else %>
	  <%'20030909 IF Request("TruckerFlag")<>1 Then %>
	<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
	  '   IF Request("TruckerFlag")<>1 AND Request("compFlag") Then 
	     IF Request("compFlag") Then %>
	       <INPUT type=button value="更新" onClick="GoReEntry()">
	  <% End If %>
	  <% IF UCase(Session.Contents("userid"))=CMPcd(0) Then %>
	       <INPUT type=hidden name=WkCNo value="<%=Request("WkCNo")%>" >
	       <INPUT type=button value="削除" onClick="GoDell()">
	  <% End If %>
	<% End If %>
	       <INPUT type=button value="キャンセル" onClick="window.close()">
	       <P>
	       <INPUT type=button value="コンテナ情報" onClick="GoConInfo()">
<% End If %>
<%end if%>
<!--2017/02/06 T.Okui Del End-->
	    </TD></TR>
	</TABLE>
</TD>
</TR>
</TABLE>
</FORM>
<!-------------画面終わり--------------------------->
</BODY></HTML>
