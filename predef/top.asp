<%
	'**********************************************
	'  【プログラムＩＤ】　: 
	'  【プログラム名称】　: 
	'
	'  （変更履歴）
	'**********************************************
	Option Explicit
	Response.Expires = 0
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
	Response.AddHeader "Pragma", "no-cache" 
%>
<%
	'**********************************************
	'共通の前提処理
	'共通関数  (Commonfunc.inc)
%>

<!--#include File="Common.inc"-->
<!--#include File="../download/download.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->

<%
	'**********************************************

	'セッションの有効性をチェック
	  CheckLoginH

	'ユーザデータ所得
	dim USER, COMPcd  			
	dim v_GamenMode
	dim v_DataCnt1
	dim v_DataCnt2
	
	'空搬出事前情報のテーブル
	dim HeaderTbl1
	dim Num
	dim v_SortKey1
	dim strWhere1
	dim strOrder1
	dim FieldName1
	dim chkAnsNo()	
	dim ObjRS,ObjConn
	dim blnSorted
	
	'空搬出事前情報のテーブル	
	dim HeaderTbl2
	dim Num2
	dim v_SortKey2	
	dim strOrder2
	dim FieldName2	
	dim chkAns2No()
	dim ObjRS2,ObjConn2
	dim TruckerSubCode
	dim blnSorted2
	
	dim wk
	dim i,x
	dim v_ItemName
	dim v_loop
	dim v_work
	dim file1,gerrmsg 
	dim Arr_SCACCode()
	dim v_SCACCode
	dim v_InOutFlag
	dim strSort
	dim abspage, pagecnt,reccnt	
	
	'201004/27 Add-S Tanaka
	v_InOutFlag=""
	'201004/27 Add-E Tanaka

	'INIファイルより設定値を取得
	dim param(2),calcDate,calcDate1
	getIni param
  	calcDate = DateAdd("d", "-"&param(2), Date)
	calcDate1 = DateAdd("d", "-"&param(1), Date)

	ReDim FieldName1(12) '2010/04/27-3 Upd C.Pestano
	'2010/05/12 Upd-S Tanaka SQLエラーになるので修正
	'FieldName1=Array("InPutDate","SenderCode","Flag2","SPB.BookNo","NumCount","ContSize1","ContType1","ContHeight1","ContMaterial1","ShipLine","FullName","TruckerCode","TruckerFlag") '2010/04/27-3 Upd C.Pestano SPB.BookNo 		
	FieldName1=Array("InPutDate","SenderCode","Flag2","SPB.BookNo","NumCount","ContSize1","ContType1","ContHeight1","ContMaterial1","SPB.ShipLine","mV.FullName","TruckerCode","TruckerFlag") '2010/04/27-3 Upd C.Pestano SPB.BookNo 		
	'2010/05/12 Upd-S Tanaka			
	ReDim FieldName2(14)	
	'2010/05/13 Upd-S Tanaka
	''2010/05/12 Upd-S C.Pestano
	''FieldName2=Array("WorkDate","Code1","Flag2","WkNo","BLContNo","ShipLine","ShipName","ContSize","ReceiveFrom","CY","DelPermitDate","FreeTime","CYCut")    
	'FieldName2=Array("WorkDate","Code1","Flag2","WkNo","BLContNo","ShipLine","ShipName","ContSize","ReceiveFrom","CY","DelPermitDate","FreeTime","CYCut","TruckerCode","TruckerFlag")    
	''2010/05/12 Upd-E C.Pestano
	FieldName2=Array("WorkDate","Code1","Flag2","WkNo","BLContNo","ShipLine","ShipName","ContSize","ReceiveFrom","CY","DelPermitDate","FreeTime","CYCut","Code2","Flag1")    
	'2010/05/13 Upd-E Tanaka

	Const CONST_ASC = "<BR><IMG border=0 src=Image/ascending.gif>"
	Const CONST_DESC = "<BR><IMG border=0 src=Image/descending.gif>"
	const gcPage = 10
	
	blnSorted = False
	blnSorted2 = False
	USER   = UCase(Session.Contents("userid"))
	COMPcd = Session.Contents("COMPcd")  	
	
	'----------------------------------------
    ' 再描画前の項目取得
   	'----------------------------------------			
	call LfGetRequestItem		
	call LfgetSCAC("",Arr_SCACCode)
	
	strSort = request.cookies("SortTbl1")
		
	if strSort <> "" then			
		if Mid(strSort,1,2) = "XX" then
			Session("Key1") = ""
		else			
			Session("Key1") = Mid(strSort,1,2)
		end if
		
		if Mid(strSort,3,1) = "0" then
			Session("KeySort1") = "ASC"
		else
			Session("KeySort1") = "DESC"
		end if	
		
		
		if Mid(strSort,4,2) = "XX" then
			Session("Key2") = ""
		else
			Session("Key2") = Mid(strSort,4,2)	'2010/04/27-3 Upd C.Pestano SPB.BookNo
		end if
		
		if Mid(strSort,6,1) = "0" then
			Session("KeySort2") = "ASC"
		else
			Session("KeySort2") = "DESC"
		end if			
		
		if Mid(strSort,7,2) = "XX" then
			Session("Key3") = ""
		else
			Session("Key3") = Mid(strSort,7,2) '2010/04/27-3 Upd C.Pestano SPB.BookNo
		end if
				
		if Mid(strSort,9,1) = "0" then
			Session("KeySort3") = "ASC"
		else
			Session("KeySort3") = "DESC"
		end if	
		'2010/04/25 Del-S Tanaka cookiesの設定が無い場合の処理を下段に移動
		'else
		'	Session("Key1") = "0"
		'	Session("KeySort1") = "ASC"
		'2010/04/25 Del-E Tanaka
	end if
	'2010/04/25 Upd-S Tanaka cookiesの設定が無い場合の処理を追加
	'strOrder1 = getSort1(Session("Key1"),Session("KeySort1"),"")
	'strOrder1 = getSort1(Session("Key2"),Session("KeySort2"),strOrder1)
	'strOrder1 = getSort1(Session("Key3"),Session("KeySort3"),strOrder1)

	if strSort <> "" then
		strOrder1 = getSort1(Session("Key1"),Session("KeySort1"),"")
		strOrder1 = getSort1(Session("Key2"),Session("KeySort2"),strOrder1)
		strOrder1 = getSort1(Session("Key3"),Session("KeySort3"),strOrder1)		
	Else
		strOrder1=" ORDER BY VBA.MAXDATE DESC,SPB.InputDate DESC, SPB.BookNo ASC "
	End If
	strSort = request.cookies("SortTbl2")	
	
	if strSort <> "" then
		if Mid(strSort,1,2) = "XX" then
			Session("TB2Key1") = ""
		else			
			Session("TB2Key1") = Mid(strSort,1,2)
		end if
		
		if Mid(strSort,3,1) = "0" then
			Session("TB2KeySort1") = "ASC"
		else
			Session("TB2KeySort1") = "DESC"
		end if	
		
		
		if Mid(strSort,4,2) = "XX" then
			Session("TB2Key2") = ""
		else
			Session("TB2Key2") = Mid(strSort,4,2) '2010/04/27-3 Upd C.Pestano SPB.BookNo
		end if
		
		if Mid(strSort,6,1) = "0" then
			Session("TB2KeySort2") = "ASC"
		else
			Session("TB2KeySort2") = "DESC"
		end if			
		
		if Mid(strSort,7,2) = "XX" then
			Session("TB2Key3") = ""
		else
			Session("TB2Key3") = Mid(strSort,7,2) '2010/04/27-3 Upd C.Pestano SPB.BookNo
		end if
				
		if Mid(strSort,9,1) = "0" then
			Session("TB2KeySort3") = "ASC"
		else
			Session("TB2KeySort3") = "DESC"
		end if
		'2010/04/25 Del-S Tanaka cookiesの設定が無い場合の処理を下段に移動
		'else
		'	Session("TB2Key1") = "0"
		'	Session("TB2KeySort1") = "ASC"
		'2010/04/25 Del-E Tanaka
	end if
	
	'2010/04/25 Upd-S Tanaka cookiesの設定が無い場合の処理を追加
	'strOrder2 = getSort2(Session("TB2Key1"),Session("TB2KeySort1"),"")
	'strOrder2 = getSort2(Session("TB2Key2"),Session("TB2KeySort2"),strOrder2)
	'strOrder2 = getSort2(Session("TB2Key3"),Session("TB2KeySort3"),strOrder2)
	'cookiesに値が存在する場合
	if strSort <> "" then
		strOrder2 = getSort2(Session("TB2Key1"),Session("TB2KeySort1"),"")
		strOrder2 = getSort2(Session("TB2Key2"),Session("TB2KeySort2"),strOrder2)
		strOrder2 = getSort2(Session("TB2Key3"),Session("TB2KeySort3"),strOrder2)
	Else
		strOrder2=" ORDER BY WorkDate_Sort ,InputDate  "
	End If
	'2010/04/25 Upd-E Tanaka	
	if v_SCACCode <> "" then
		strWhere1 = strWhere1 & " AND SPB.ShipLine = '" & Trim(v_SCACCode) & "'"
	else
		strWhere1 = ""		
	end if
	
	Call getDataTbl1(strWhere1)
	Call getDataTbl2()
	
	if v_GamenMode = "U" then		
		call LfUpdTruckerAns()
		Response.redirect "./top.asp?pagenum=" & CInt(Request("pagenum")) & "&pagenum2=" & CInt(Request("pagenum2"))
	end if	
	
	if v_GamenMode = "P" then		
		wReportName="搬入票" 
		wReportID="dmo320" 		
		wOutFileName=gfReceiveReportMultiple()				
		file1	= server.mappath(gOutFileForder & wOutFileName)
		if not gfdownloadFile(file1, wOutFileName) then
			wMsg = Replace(gerrmsg,"<br>","\n")
		end if		
	end if

    'Y.TAKAKUWA Add-S 2013-02-18
    Session("InOutFlag") = v_InOutFlag
    'Y.TAKAKUWA Add-E 2013-02-18
    
	Function LfGetRequestItem()
		If Request.form("Gamen_Mode")<>"R" Then
			v_GamenMode = Request.form("Gamen_Mode") 
			v_SCACCode = Request.form("cmbSCACCode")
			'2010/04/26 Add-S Tanaka ページ移動の場合のみページ移動時の船社を再設定する
			if trim(v_SCACCode)="" and Trim(Request.form("SCACSrhFlag")) = "" Then
				v_SCACCode=Request("SCACCode")
			End If
			v_InOutFlag = Request.form("cmbINOut")
			if trim(v_InOutFlag)="" and Trim(Request.form("InOutSrhFlag")) = "" Then
				v_InOutFlag=Request("InOutF")
			End If
			'2010/04/26 Add-E Tanaka
			v_DataCnt1 = Request.form("DataCnt1") 
			v_DataCnt2 = Request.form("DataCnt2") 
		Else
			'2010/05/06 Add-S C.Pestano
			v_SCACCode = Request.form("cmbSCACCode")
			v_InOutFlag = Request.form("cmbINOut")
			'2010/05/06 Add-E C.Pestano
		End If
	End Function

Function getDataTbl1(strWhere)
	dim StrSQL		
  	dim x
	dim ctr
	dim i
	
	Num = 0
	
	On Error Resume Next	
	
	ConnDBH ObjConn, ObjRS
		
	WriteLogH "b101", "空搬出事前情報一覧", "04",""
	
	ReDim HeaderTbl1(13)
	
	HeaderTbl1(0) = "入力日"
	HeaderTbl1(1) = "指示元"
	HeaderTbl1(2) = "指示元<BR>へ回答"	
	HeaderTbl1(3) = "回答"
	HeaderTbl1(4) = "ブッキング番号"
	HeaderTbl1(5) = "ピック済"
	HeaderTbl1(6) = "SZ"
	HeaderTbl1(7) = "タイプ"
	HeaderTbl1(8) = "H"
	HeaderTbl1(9) = "材質"
	HeaderTbl1(10) = "船社"
	HeaderTbl1(11) = "船名"
	HeaderTbl1(12) = "指示先"
	HeaderTbl1(13) = "指示先<BR>回答"
		 
	for ctr = 1 to 3	
		Session(CSTR("Key" & ctr))	
		if Session(CSTR("Key" & ctr)) <> "" then
			'2010/04/27-3 Upd-S C.Pestano
			Select Case Session(CSTR("Key" & ctr))				
				Case "00" '入力日
					HeaderTbl1(0) = HeaderTbl1(0) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "01" '指示元
					HeaderTbl1(1) = HeaderTbl1(1) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "02" '指示元へ回答
					HeaderTbl1(2) = HeaderTbl1(2) & getImage(Session(CSTR("KeySort" & ctr)))
					blnSorted = True				
				Case "03" 'ブッキング番号
					HeaderTbl1(4) = HeaderTbl1(4) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "04" 'ピック済
					HeaderTbl1(5) = HeaderTbl1(5) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "05" 'サイズ
					HeaderTbl1(6) = HeaderTbl1(6) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "06" 'タイプ
					HeaderTbl1(7) = HeaderTbl1(7) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "07" '高さ
					HeaderTbl1(8) = HeaderTbl1(8) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "08" '材質
					HeaderTbl1(9) = HeaderTbl1(9) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "09" '船社
					HeaderTbl1(10) = HeaderTbl1(10) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "10" '船名
					HeaderTbl1(11) = HeaderTbl1(11) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "11" '指示先
					HeaderTbl1(12) = HeaderTbl1(12) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "12" '指示先へ回答
					HeaderTbl1(13) = HeaderTbl1(13) & getImage(Session(CSTR("KeySort" & ctr)))	
					blnSorted = True							
			  End Select
			  '2010/04/27-3 Upd-E C.Pestano
		end if	  
	next	
		
	'" LEFT JOIN Booking as BOK on VECO.VslCode=BOK.VslCode and VECO.VoyCtrl=BOK.VoyCtrl and VECO.BookNo=BOK.BookNo"&_
	'" AND DateDiff(day,SPB.InputDate,'"&calcDate&"')<=0" &_	
	'" IsNull(CASE (VPC.Picks) WHEN '1' THEN VPC.PickPlace ELSE '複数' END ,'') PickPlace,"&_
	'" LEFT JOIN ViewPickupCnt AS VPC ON BOK.VslCode = VPC.VslCode AND BOK.VoyCtrl = VPC.VoyCtrl AND BOK.BookNo = VPC.BookNo"&_ ,"&_
	'	StrSQL = "SELECT * FROM(SELECT DISTINCT SPB.BookNo, mV.FullName, SPB.SenderCode," &_
	'			 " ISNULL(CONVERT(varchar(10),SPB.InputDate,111),'') as InputDate, SPB.TruckerCode, SPB.TruckerFlag,"&_
	'			 " (CASE SPB.TruckerFlag WHEN 0 THEN '未' WHEN 1 THEN 'Yes' WHEN 2 THEN 'No' ELSE ' ' END) as Flag2,"&_ 
	'			 " SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1, mS.OpeCode,"&_
	'			 " (CASE WHEN mU.UserType = '5' THEN mU.HeadCompanyCode ELSE SPB.SenderCode END) as Code1, "&_
	'			 " (CASE WHEN mU.UserType = '5' THEN mU.TTName ELSE mU.TTName END) as TruckerName, "&_
	'			 " IsNull(VEC.numC,'') as NumCount,"&_
	'			 " SPB.Comment1,SPB.Comment2, mU.HeadCompanyCode, mU.UserType ,SPB.ShipLine,VSLS.CYCut "&_
	'			 " FROM BookingAssign AS SPB "&_
	'			 " LEFT JOIN Booking As BOK ON SPB.BookNo = BOK.BookNo"&_
	'			 " LEFT JOIN ViewExportCnt As VEC ON SPB.BookNo = VEC.BookNo"&_			 			 
	'			 " LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_			 			 
	'			 " LEFT JOIN VslSchedule AS VSLS ON BOK.VslCode = VSLS.VslCode AND BOK.VoyCtrl = VSLS.VoyCtrl"&_ 
	'			 " LEFT JOIN mVessel AS mV ON BOK.VslCode = mV.VslCode"&_
	'			 " LEFT JOIN mShipLine AS mS ON SPB.ShipLine = mS.ShipLine"&_ 			 
	'			 " WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R'"&_
	'			 " AND '"&calcDate&"' <= SPB.InputDate" &_
	'			 " ) AS T"&_
	'			 " WHERE T.Flag2='未' OR T.Flag2=' ' "&_			 			 
	'			 strWhere & strOrder1		
	
	'2020/06/22 Upd-S Fujiyama 空バン１行目が削除で表示しなくなる不具合対策
	'	StrSQL = "SELECT ISNULL(CONVERT(varchar(10),SPB.InputDate,111),'') as InputDate,"&_
	'			 " (CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE SPB.SenderCode END )as SenderCode,"&_
	'			 " SPB.SenderCode as SenderCode1, "&_
	'			 " (CASE SPB.TruckerFlag WHEN 0 THEN '未' WHEN 1 THEN 'Yes' WHEN 2 THEN 'No' ELSE ' ' END) as Flag2,"&_
	'			 " SPB.BookNo, "&_
	'			 " IsNull(VEC.numC,'') as NumCount,"&_
	'			 " SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1,"&_
	'			 " SPB.ShipLine,"&_
	'			 " mV.FullName, "&_
	'			 " SPB.TruckerCode,"&_
	'			 " SPB.TruckerFlag"&_
	'   		 " FROM BookingAssign AS SPB "&_
	'			 " INNER JOIN (select BookNo, SenderCode, TruckerCode, MIN(Seq) AS Seq from BookingAssign Group BY BookNo, SenderCode, TruckerCode) AS BA "&_
	'			 "       ON SPB.BookNo = BA.BookNo AND SPB.SenderCode = BA.SenderCode AND SPB.TruckerCode = BA.TruckerCode AND SPB.Seq = BA.Seq " &_
	'			 " LEFT JOIN Booking As BOK ON SPB.BookNo = BOK.BookNo"&_
	'			 " LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode"&_
	'			 " LEFT JOIN ViewExportCnt As VEC ON BOK.VslCode = VEC.VslCode AND BOK.VoyCtrl = VEC.VoyCtrl AND BOK.BookNo = VEC.BookNo AND SPB.SenderCode = VEC.SenderCode AND SPB.TruckerCode = VEC.TruckerCode "&_
	'			 " LEFT JOIN mVessel AS mV ON BOK.VslCode = mV.VslCode"&_
	'			 " LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
	'			 " WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R'"&_
	'			 " AND '"&calcDate&"' <= SPB.InputDate "&_
	'			 strWhere & strOrder1

	StrSQL = "SELECT ISNULL(CONVERT(varchar(10),SPB.InputDate,111),'') as InputDate,"&_
			 " (CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE SPB.SenderCode END )as SenderCode,"&_
			 " SPB.SenderCode as SenderCode1, "&_
			 " (CASE SPB.TruckerFlag WHEN 0 THEN '未' WHEN 1 THEN 'Yes' WHEN 2 THEN 'No' ELSE ' ' END) as Flag2,"&_
			 " SPB.BookNo, "&_
			 " IsNull(VEC.numC,'') as NumCount,"&_
			 " SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1,"&_
			 " SPB.ShipLine,"&_
			 " mV.FullName, "&_
			 " SPB.TruckerCode,"&_
			 " SPB.TruckerFlag"&_
    		 " FROM BookingAssign AS SPB "&_
			 " INNER JOIN (select BookNo, SenderCode, TruckerCode, MIN(Seq) AS Seq from BookingAssign where Process <> 'D' Group BY BookNo, SenderCode, TruckerCode) AS BA "&_
			 "       ON SPB.BookNo = BA.BookNo AND SPB.SenderCode = BA.SenderCode AND SPB.TruckerCode = BA.TruckerCode AND SPB.Seq = BA.Seq " &_
			 " LEFT JOIN Booking As BOK ON SPB.BookNo = BOK.BookNo"&_
			 " LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode"&_
			 " LEFT JOIN ViewExportCnt As VEC ON BOK.VslCode = VEC.VslCode AND BOK.VoyCtrl = VEC.VoyCtrl AND BOK.BookNo = VEC.BookNo AND SPB.SenderCode = VEC.SenderCode AND SPB.TruckerCode = VEC.TruckerCode "&_
			 " LEFT JOIN mVessel AS mV ON BOK.VslCode = mV.VslCode"&_
			 " LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
			 " WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R'"&_
			 " AND '"&calcDate&"' <= SPB.InputDate "&_
			 strWhere & strOrder1
	'2020/06/22 Upd-E Fujiyama
	
	ObjRS.PageSize = 100	
	ObjRS.CacheSize = 100
	ObjRS.CursorLocation = 3
	ObjRS.Open StrSQL, ObjConn
	Num = ObjRS.recordcount			

	if CInt(Num) > 100 then
		If CInt(Request("pagenum")) = 0 Then
			ObjRS.AbsolutePage = 1
		Else
			If CInt(Request("pagenum")) <= ObjRS.PageCount Then
				ObjRS.AbsolutePage = CInt(Request("pagenum"))				
			Else
				ObjRS.AbsolutePage = 1				
			End If			
		End If		
	End If	
	
	if err <> 0 then			
		DisConnDBH ObjConn, ObjRS	'DB切断
		jampErrerP "2","b301","01","空搬出事前情報一覧","102","SQL:<BR>" & strSQL & err.description
		Exit Function
	end if
		
  'エラートラップ解除
  on error goto 0   
End Function	

Function getDataTbl2()
	dim StrSQL		
  	dim x
	dim ctr
	dim tmptime
	'2010/04/27 Add-S Tanaka
	dim strInOutWhere
	'2010/04/27 Add-E Tanaka
	On Error Resume Next

	'2010/04/27 Add-S Tanaka
	If Trim(v_InOutFlag)="1" Then
		strInOutWhere =" WHERE Type='FULLOUT' "
	ElseIf Trim(v_InOutFlag)="2" Then
		strInOutWhere =" WHERE Type='FULLIN' "
	Else
		strInOutWhere =" "
	End If

	
	'2010/04/27 Add-E Tanaka
	ConnDBH ObjConn2, ObjRS2	
	
	WriteLogH "b401", "実搬入事前情報一覧", "01", ""
	
	ReDim HeaderTbl2(18)	'Y.TAKAKUWA Upd 16
	
	HeaderTbl2(0) = "搬入票<BR>出力"	
	HeaderTbl2(1) = "搬出入<BR>予定日"
	HeaderTbl2(2) = "指示元"
	HeaderTbl2(3) = "指示元<BR>へ回答"
	HeaderTbl2(4) = "回答"	
	HeaderTbl2(5) = "作業<BR>番号"
	HeaderTbl2(6) = "コンテナ番号<BR>/BL番号"
	HeaderTbl2(7) = "船社"
	HeaderTbl2(8) = "船名"
	HeaderTbl2(9) = "SZ"
	HeaderTbl2(10) = "搬入元"
	HeaderTbl2(11) = "CY"
	HeaderTbl2(12) = "搬出許可日"
	HeaderTbl2(13) = "フリー<BR>タイム"
	HeaderTbl2(14) = "CYカット日"					
	'2010/05/12 Upd-S C.Pestano		
	HeaderTbl2(15) = "指示先"
	HeaderTbl2(16) = "指示先<BR>回答"
	'2010/05/12 Upd-E C.Pestano		
	'2013-02-18 Y.TAKAKUWA Add-S
	HeaderTbl2(17) = "ロック<BR/>指示先"
	HeaderTbl2(18) = "ロック中<BR/>ヘッドID"
	'2013-02-18 Y.TAKAKUWA Add-E
	
	for ctr = 1 to 3	
		Session(CSTR("TB2Key" & ctr))	
		if Session(CSTR("TB2Key" & ctr)) <> "" then
			'2010/04/27-3 Upd-S C.Pestano
			Select Case Session(CSTR("TB2Key" & ctr))
				Case "00" '搬出入予定日
					HeaderTbl2(1) = HeaderTbl2(1) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "01" '指示元
					HeaderTbl2(2) = HeaderTbl2(2) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "02" '指示元へ回答
					HeaderTbl2(3) = HeaderTbl2(3) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "03" '作業番号
					HeaderTbl2(5) = HeaderTbl2(5) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "04" 'コンテナ番号/BL番号
					HeaderTbl2(6) = HeaderTbl2(6) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "05" '船社
					HeaderTbl2(7) = HeaderTbl2(7) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "06" '船名
					HeaderTbl2(8) = HeaderTbl2(8) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "07" 'SZ
					HeaderTbl2(9) = HeaderTbl2(9) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "08" '搬入元
					HeaderTbl2(10) = HeaderTbl2(10) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "09" 'CY
					HeaderTbl2(11) = HeaderTbl2(11) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "10" '搬出可否
					HeaderTbl2(12) = HeaderTbl2(12) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "11" 'フリータイム
					HeaderTbl2(13) = HeaderTbl2(13) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "12" 'CYカット日
					HeaderTbl2(14) = HeaderTbl2(14) & getImage(Session(CSTR("TB2KeySort" & ctr)))			
				'2010/05/12 Upd-S C.Pestano
				Case "13" '指示先
					HeaderTbl2(15) = HeaderTbl2(15) & getImage(Session(CSTR("TB2KeySort" & ctr)))			
				Case "14" '指示先<BR>回答
					HeaderTbl2(16) = HeaderTbl2(16) & getImage(Session(CSTR("TB2KeySort" & ctr)))			
				'2010/05/12 Upd-E C.Pestano
			End Select
			'2010/04/27-3 Upd-E C.Pestano
		end if	  
	next					
	
	'2010/04/27-3 Upd-S C.Pestano
	StrSQL ="SELECT * FROM (SELECT T.* FROM (SELECT 'FULLOUT' As Type,ITC.DeliverTo,ITC.BLNo, " & _
			"ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate, "&_
			"CASE WHEN ITC.WorkDate Is NULL THEN '9999/12/31' WHEN ITC.WorkDate ='' THEN '9999/12/31' ELSE ITC.WorkDate END as WorkDate_Sort, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
			"ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
			"END) as Code1, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			"ELSE ITC.TruckerSubName1 "&_
			"END) as Name1, "&_
			"(CASE "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
			"ELSE Null END) "&_
			"WHEN 0 THEN '未' "&_
			"WHEN 1 THEN 'Yes' "&_
			"WHEN 2 THEN 'No' "&_
			"ELSE ' ' END) as Flag2, "&_
			"ITC.WkNo, "&_
			"ITC.FullOutType, "&_
			"ITC.BLNo as BLContNo, A.ShipLine, A.FullName as ShipName, ''  as ContSize, "&_
			"SUBSTRING(A.RecTerminal,1,2) as CY, "&_
			"ISNULL(CONVERT(varchar(10),A.FreeTime,111),'') as FreeTime, "&_			
			"CASE WHEN A.FreeTime Is NULL THEN '9999/12/31' WHEN A.FreeTime ='' THEN '9999/12/31' ELSE A.FreeTime END as FreeTime_Sort, "&_			
			"ITC.DeliverTo1,ITC.WorkCompleteDate, "&_
			"ITC.ReturnDateStr, "&_
			"(CASE WHEN ITC.FullOutType = '1' THEN (CASE WHEN INC.ReturnTime IS NULL THEN '未' ELSE '済' END) "&_
			"ELSE ' ' END) as ReturnValue, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) as Code2, "&_
			"(CASE WHEN "&_
			"(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) IS NULL THEN ' ' ELSE "&_
			"(CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITR.TruckerFlag1 "&_
			"END) "&_
			"WHEN '0' THEN '未' "&_
			"WHEN '1' THEN 'Yes' "&_
			"ELSE 'No' END) "&_
			"END) as Flag1, "&_
			"ITC.Comment1, ITC.Comment2, "&_
			"ITC.ReturnDateVal, ITC.UpdtUserCode, "&_
			"ITR.TruckerFlag1,ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
			"mU.HeadCompanyCode, mU.UserType, "&_
			"A.CYDelTime, INC.ReturnTime,ITC.InputDate, "&_
			"'' as InputDate1,'' as WorkDate1,'' as BookNo,'' as ContNo,'' as VslName,'' as ContHeight,'' as TareWeight, "&_
			"'' as ReceiveFrom,ISNULL(CONVERT(varchar(10),INC.DelPermitDate,111),'-') as DelPermitDate,"&_
			"CASE WHEN INC.DelPermitDate Is NULL THEN '9999/12/31' WHEN INC.DelPermitDate ='' THEN '9999/12/31' ELSE INC.DelPermitDate END as DelPermitDate_Sort, "&_			
			"'' as CYCut,'9999/12/31' as CYCut_Sort,'' as CYCut1,'' as WorkComplete,'' as WorkComplete1, "&_
			"ITC.TruckerSubCode1,ITC.TruckerSubCode2,ITC.TruckerSubCode3,ITC.TruckerSubCode4,ITC.WkContrlNo, "&_
			"'' as Nine , '' as Comment3,'' as RegisterCode "&_	

            " ,'' As LOID "&_
			" ,'' AS LoHeadID "&_
			" ,'' AS LoDriverName "&_
			" ,'' AS GroupFlag "&_
					
			"FROM hITCommonInfo ITC "&_
			"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
			"LEFT JOIN ImportCont AS INC ON (ITC.ContNo=INC.ContNo) "&_
			"LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode) "&_
			"LEFT JOIN (SELECT Distinct BL.BLNo,INC.FreeTime, "&_
			"MIN(INC.CYDelTime) AS CYDelTime,mV.ShipLine, mV.FullName, BL.RecTerminal "&_
		"FROM ImportCont AS INC "&_
			"LEFT JOIN mVessel AS mV ON INC.VslCode = mV.VslCode "&_
		"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
		
			"GROUP BY BL.BLNo,INC.FreeTime,mV.ShipLine, mV.FullName, BL.RecTerminal) A ON  A.BLNo=ITC.BLNo "&_
			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType='4' AND (ITC.RegisterCode='"& USER &"' "&_
		"OR ITC.TruckerSubCode1='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode2='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode3='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode4='"& COMPcd &"') "&_
			"AND (ITC.WorkCompleteDate IS Null) "&_
		"UNION ALL "&_
			"SELECT DISTINCT 'FULLOUT' As Type,ITC.DeliverTo,ITC.BLNo, "&_
		"ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate,"&_
		"CASE WHEN ITC.WorkDate Is NULL THEN '9999/12/31' WHEN ITC.WorkDate ='' THEN '9999/12/31' ELSE ITC.WorkDate END as WorkDate_Sort, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
			"ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
		"END) as Code1, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			"ELSE ITC.TruckerSubName1 "&_
			"END) as Name1, "&_
			"(CASE "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
			"ELSE Null END) "&_
			"WHEN 0 THEN '未' "&_
			"WHEN 1 THEN 'Yes' "&_
			"WHEN 2 THEN 'No' "&_
			"ELSE ' ' END) as Flag2, "&_
			"ITC.WkNo, "&_
			"ITC.FullOutType, "&_
			"ITC.ContNo as BLContNo, "&_
			"mV.ShipLine, "&_
			"mV.FullName as ShipName, "&_
			"CON.ContSize, "&_
			"SUBSTRING(BL.RecTerminal,1,2) as CY, "&_
			"ISNULL(CONVERT(varchar(10),INC.FreeTime,111),'') as FreeTime, "&_			
			"CASE WHEN INC.FreeTime Is NULL THEN '9999/12/31' WHEN INC.FreeTime ='' THEN '9999/12/31' ELSE INC.FreeTime END as FreeTime_Sort,"&_			
			"ITC.DeliverTo1, "&_
			"ITC.WorkCompleteDate, "&_
			"ITC.ReturnDateStr, "&_
			"(CASE WHEN ITC.FullOutType = '1' THEN (CASE WHEN INC.ReturnTime IS NULL THEN '未' "&_
			" ELSE '済' END) "&_
			"ELSE ' ' END) as ReturnValue, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) as Code2, "&_
			"(CASE WHEN "&_
			"(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) IS NULL THEN ' ' "&_
			"ELSE "&_
			"(CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITR.TruckerFlag1 "&_
			"END) "&_
			"WHEN '0' THEN '未' WHEN '1' THEN 'Yes' ELSE 'No' END) END) as Flag1, "&_
			"ITC.Comment1, ITC.Comment2, "&_
			"ITC.ReturnDateVal, ITC.UpdtUserCode, "&_
			"ITR.TruckerFlag1,ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
			"mU.HeadCompanyCode, mU.UserType, "&_
			"INC.CYDelTime, INC.ReturnTime,ITC.InputDate, "&_
			"'' as InputDate1,'' as WorkDate1,'' as BookNo,'' as ContNo,'' as VslName,'' as ContHeight,'' as TareWeight, "&_
			"'' as ReceiveFrom,ISNULL(CONVERT(varchar(10),INC.DelPermitDate,111),'-') as DelPermitDate,"&_
			"CASE WHEN INC.DelPermitDate Is NULL THEN '9999/12/31' WHEN INC.DelPermitDate ='' THEN '9999/12/31' ELSE INC.DelPermitDate END as DelPermitDate_Sort, "&_
			"'' as CYCut,'9999/12/31' as CYCut_Sort,'' as CYCut1,'' as WorkComplete,'' as WorkComplete1, "&_
			"ITC.TruckerSubCode1,ITC.TruckerSubCode2,ITC.TruckerSubCode3,ITC.TruckerSubCode4,ITC.WkContrlNo, "&_
			"'' as Nine , '' as Comment3,'' as RegisterCode "&_	
			
			" ,LOInfo.InputID As LOID "&_
			" ,LOInfo.LoHeadID AS LoHeadID "&_
			" ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_
            "  LOInfo.InputID "&_
            "  ELSE "&_
            "  LODriver.LoDriverName "&_
            "  END AS LoDriverName "&_
            
             "  ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_ 
            "    CASE WHEN LOInfo.LoStatus <> 2 THEN "&_
            "      CASE WHEN LoOwnGroup.LoGroupID IS NULL THEN '0' "&_
            "      ELSE '2' END "&_
            "    ELSE '2' END "&_
            "  ELSE '1' END AS  "&_
            "  GroupFlag "&_
			
			"FROM hITCommonInfo ITC "&_
		"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
			"LEFT JOIN ImportCont AS INC ON ITC.ContNo=INC.ContNo "&_
			"LEFT JOIN mVessel AS mV On INC.VslCode=mV.VslCode "&_
			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
			"LEFT JOIN Container AS CON ON INC.VslCode = CON.VslCode AND INC.VoyCtrl=CON.VoyCtrl AND INC.ContNo=CON.ContNo "&_
			
			"LEFT JOIN ("&_
                        "SELECT LoLockOnInfo.WkNo, "&_
                                 "LoLockOnInfo.ContNo, "&_
                                 "LoLockOnInfo.SeqNumber, "&_
                                 "LoLockOnInfo.LoStatus, "&_
                                 "LoLockOnInfo.KaccsSend, "&_
                                 "LoLockOnInfo.WorkDate, "&_
                                 "LoLockOnInfo.HiTSUserID, "&_
                                 "CASE WHEN ISNULL(LoLockOnInfo.LoDriverID,'') <> '' THEN LoLockOnInfo.LoDriverID "&_ 
                                      "ELSE LoLockOnInfo.InputID "&_
                                 "END AS InputID, "&_
                                 "LoLockOnInfo.LoDriverID, "&_
                                 "LoLockOnInfo.LoHeadID "&_      
                        "FROM LoLockOnInfo "&_ 
                        
	        ") AS LOInfo ON ITC.WkNo = LOInfo.WkNo AND ITC.ContNo = LOInfo.ContNo AND LOInfo.LoStatus<>'c' AND LOInfo.LoStatus<>'d' "&_
		    
		    
		    "        LEFT JOIN LomDriver AS LODriver ON LOInfo.InputID = "&_
			"        CASE WHEN(LEN(LOInfo.InputID) = 2) THEN LODriver.LoDriverHeadID "&_
			"             WHEN(LEN(LOInfo.InputID)>=6 AND LEN(LOInfo.InputID)<=12) THEN LODriver.LoDriverID "&_
			"             ELSE LODriver.LoDriverID "&_
			"        END "&_
			"        LEFT JOIN LomGroup ON LOInfo.InputID = LomGroup.LoGroupID AND LomGroup.HiTSUserID='" & USER & "'" &_
			
			"        OUTER APPLY "&_
            "        ( "&_
            "        SELECT TOP 1 * FROM "&_
            "        LoGroupeDriver  "&_
            "        WHERE LoGroupeDriver.LoGroupID = LOInfo.InputID AND LoGroupeDriver.HiTSUserID='" & USER & "' "&_
            "        ) LoGroupeDriver "&_
			
			"        OUTER APPLY "&_
            "        ( "&_
            "        SELECT TOP 1 LoGroupeDriver.* FROM LomGroup "&_
            "        INNER JOIN LoGroupeDriver ON  LomGroup.LoGroupID=LoGroupeDriver.LoGroupID AND LomGroup.HiTSUserID = LoGroupeDriver.HiTSUserID "&_
            "        WHERE LomGroup.HiTSUserID='" & USER & "' AND LomGroup.LoGroupID=LOInfo.InputID "&_
            "        ) LoOwnGroup "&_
			
			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType<>'4' AND (ITC.RegisterCode='"& USER &"' "&_
			"OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"')" &_
			"AND ( ITC.WorkCompleteDate IS Null) "&_
			") AS T " &_
            "UNION ALL " &_			
			"SELECT T.* FROM (SELECT DISTINCT 'FULLIN' As Type,'' as DeliverTo, '' as BLNo, ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate," &_
			"CASE WHEN ITC.WorkDate Is NULL THEN '9999/12/31' WHEN ITC.WorkDate ='' THEN '9999/12/31' ELSE ITC.WorkDate END as WorkDate_Sort, "&_
            "        (CASE " &_
    	    "            WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) " &_
            "            WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 " &_
            "            WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
            "            WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
            "            ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) " &_
		    "        END) as Code1, " &_
		    "		 (CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"			WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"			WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"			WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			"			ELSE ITC.TruckerSubName1 "&_
			"		 END) as Name1, "&_
		    "		(CASE "&_
			"		(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
			"		WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
			"		WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"		WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"		ELSE Null END) "&_
			"		WHEN 0 THEN '未' "&_
			"		WHEN 1 THEN 'Yes' "&_
			"		WHEN 2 THEN 'No' "&_
			"		ELSE ' ' END) as Flag2, "&_			
		    "       ITC.WkNo, '' as FullOutType,ITC.ContNo as BLContNo, CYV.ShipLine,CYV.VslName as ShipName, " &_
			"		CYV.ContSize,BOK.RecTerminal as CY, '' as FreeTime,'9999/12/31' as FreeTime2, '' as DeliverTo1, " &_			
		    "       '' as WorkCompleteDate,'' as ReturnDateStr,'' as ReturnValue, " &_
		    "        (CASE " &_
			"            WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
			"            WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 " &_
			"            WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
		    "            WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
		    "            ELSE ITC.TruckerSubCode1 " &_
		    "        END) as Code2, " &_
		    "        (CASE WHEN " &_
		    "              (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
		    "                    WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
		    "                    WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 " &_
		    "                    WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
		    "                    ELSE ITC.TruckerSubCode1 " &_
		    "               END) IS NULL THEN ' '" &_
		    "              ELSE " &_
		    "                    (CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 " &_
		    "                                WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 " &_
		    "                                WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 " &_
		    "                                WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
		    "                                ELSE ITR.TruckerFlag1 " &_
		    "                     END) " &_
		    "             WHEN '0' THEN '未' " &_
		    "             WHEN '1' THEN 'Yes' " &_
		    "             ELSE 'No' END) " &_
		    "          END) as Flag1, " &_
			"        SUBSTRING(ITC.Comment1,1,10) as Comment1, SUBSTRING(ITC.Comment2,1,10) as Comment2, " &_
			" 		'' as ReturnDateVal, ITC.UpdtUserCode,ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, " &_
			"		ITR.TruckerFlag4, mU.HeadCompanyCode, mU.UserType,'' as CYDelTime,'' as ReturnTime," &_
			"		CONVERT(varchar,ITC.InputDate,11) as InputDate,ITC.InputDate as InputDate1, " &_ 
			"		ITC.WorkDate as WorkDate1,CYV.BookNo,ITC.ContNo,CYV.VslName as VslName, " &_
			" 		CYV.ContHeight,CASE ISNULL(CYV.TareWeight,0) WHEN 0 THEN '-' ELSE CYV.TareWeight END TareWeight, " &_
			"		SUBSTRING(CYV.ReceiveFrom,1,20) as ReceiveFrom, '-' as DelPermitDate,'9999/12/31' as DelPermitDate_Sort," &_
			"		ISNULL(CONVERT(varchar(10),VSLS.CYCut,111),'') as CYCut,"&_
			"		CASE WHEN VSLS.CYCut Is NULL THEN '9999/12/31' WHEN VSLS.CYCut ='' THEN '9999/12/31' ELSE VSLS.CYCut END as CYCut_Sort, "&_
			" 		VSLS.CYCut as CYCut1, " &_			
			"		CONVERT(varchar,ITC.WorkCompleteDate,11) + ' ' + Substring(CONVERT(varchar,ITC.WorkCompleteDate,8),1,5) as WorkComplete, " &_
			"		ITC.WorkCompleteDate as WorkComplete1, " &_
			"		ITC.TruckerSubCode1, ITC.TruckerSubCode2, ITC.TruckerSubCode3, ITC.TruckerSubCode4,  ITC.WkContrlNo, " & _
			"          (CASE "&_
			"                WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN '1' "&_
			"                WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN '2' "&_
			"                WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN '3' "&_
			"                WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '4' "&_
			"                ELSE '0' END) as Nine, "&_		
			"        SUBSTRING(ITC.Comment3,1,10) as Comment3, " &_
		    "        ITC.RegisterCode " &_  
		    
		    " ,LOInfo.InputID As LOID "&_
			" ,LOInfo.LoHeadID AS LoHeadID "&_
			
			" ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_
            "  LOInfo.InputID "&_
            "  ELSE "&_
            "  LODriver.LoDriverName "&_
            "  END AS LoDriverName "&_
            
			"  ,CASE WHEN LEN(LOInfo.InputID)= 2 THEN "&_ 
            "    CASE WHEN LOInfo.LoStatus <> 2 THEN "&_
            "      CASE WHEN LoOwnGroup.LoGroupID IS NULL THEN '0' "&_
            "      ELSE '2' END "&_
            "    ELSE '2' END "&_
            "  ELSE '1' END AS  "&_
            "  GroupFlag "&_
		    
    	    "        FROM hITCommonInfo AS ITC " &_
		    "        INNER JOIN CYVanInfo AS CYV ON ITC.ContNo=CYV.ContNo AND ITC.WkNo = CYV.WkNo " &_
		    "        INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo " &_
		    "        INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode " &_
		    "        LEFT JOIN ExportCont AS EPC ON CYV.ContNo = EPC.ContNo AND CYV.BookNo = EPC.BookNo and CYV.VslCode = EPC.VslCode " &_
		    "        LEFT JOIN VslSchedule AS VSLS ON EPC.VslCode = VSLS.VslCode AND EPC.VoyCtrl = VSLS.VoyCtrl " &_
		    "        LEFT JOIN Booking AS BOK ON EPC.VslCode = BOK.VslCode AND EPC.VoyCtrl = BOK.VoyCtrl AND EPC.BookNo = BOK.BookNo " &_
		    
            "        LEFT JOIN ("&_
		                         "SELECT LoLockOnInfo.WkNo, "&_
                                        "LoLockOnInfo.ContNo, "&_
                                        "LoLockOnInfo.SeqNumber, "&_
                                        "LoLockOnInfo.LoStatus, "&_
                                        "LoLockOnInfo.KaccsSend, "&_
                                        "LoLockOnInfo.WorkDate, "&_
                                        "LoLockOnInfo.HiTSUserID, "&_
                                        "CASE WHEN ISNULL(LoLockOnInfo.LoDriverID,'') <> '' THEN LoLockOnInfo.LoDriverID "&_ 
                                             "ELSE LoLockOnInfo.InputID "&_
                                        "END AS InputID, "&_
                                        "LoLockOnInfo.LoDriverID, "&_
                                        "LoLockOnInfo.LoHeadID "&_      
                                 "FROM LoLockOnInfo "&_ 
		    ") AS LOInfo ON ITC.WkNo = LOInfo.WkNo AND ITC.ContNo = LOInfo.ContNo AND LOInfo.LoStatus<>'c' AND LOInfo.LoStatus<>'d' "&_
		    
		    "        LEFT JOIN LomDriver AS LODriver ON LOInfo.InputID = "&_
			"        CASE WHEN(LEN(LOInfo.InputID) = 2) THEN LODriver.LoDriverHeadID "&_
			"             WHEN(LEN(LOInfo.InputID)>=6 AND LEN(LOInfo.InputID)<=12) THEN LODriver.LoDriverID "&_
			"             ELSE LODriver.LoDriverID "&_
			"        END "&_
		    "        LEFT JOIN LomGroup ON LOInfo.InputID = LomGroup.LoGroupID AND LomGroup.HiTSUserID='" & USER & "'" &_
		    
		    "        OUTER APPLY "&_
            "        ( "&_
            "        SELECT TOP 1 * FROM "&_
            "        LoGroupeDriver  "&_
            "        WHERE LoGroupeDriver.LoGroupID = LOInfo.InputID AND LoGroupeDriver.HiTSUserID='" & USER & "' "&_
            "        ) LoGroupeDriver "&_
            
            "        OUTER APPLY "&_
            "        ( "&_
            "        SELECT TOP 1 LoGroupeDriver.* FROM LomGroup "&_
            "        INNER JOIN LoGroupeDriver ON  LomGroup.LoGroupID=LoGroupeDriver.LoGroupID AND LomGroup.HiTSUserID = LoGroupeDriver.HiTSUserID "&_
            "        WHERE LomGroup.HiTSUserID='" & USER & "' AND LomGroup.LoGroupID=LOInfo.InputID" &_
            "        ) LoOwnGroup "&_
		    
		    "        WHERE ITC.Process='R' AND ITC.WkType='3' AND (ITC.RegisterCode='"& USER &"' " &_
		    "        OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' " &_
		    "        OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"')" &_
		    "        AND (ITC.WorkCompleteDate IS Null) " &_
			"        ) AS T ) AS A " &_ 
    		strInOutWhere & strOrder2
			''2010/04/27-3 Upd-E C.Pestano
    		
	ObjRS2.PageSize = 100
	ObjRS2.CacheSize = 100
	ObjRS2.CursorLocation = 3
	ObjRS2.Open StrSQL, ObjConn2
	
	'Y.TAKAKUWA DebugADD-S 2013-02-12
	' Response.Write "<b style='color:red'>DEBUG-MODE-START</b><BR/>" 
	' Response.Write StrSQL
	' Response.Write "<BR/><b style='color:red'>DEBUG-MODE-END</b>"
	'Y.TAKAKUWA DebugADD-E 2013-02-12
	
	Num2 = ObjRS2.recordcount	
	
	if Num2 > 100 then
		If CInt(Request("pagenum2")) = 0 Then
			ObjRS2.AbsolutePage = 1
		Else
			If CInt(Request("pagenum2")) <= ObjRS2.PageCount Then
				ObjRS2.AbsolutePage = CInt(Request("pagenum2"))
			Else
				ObjRS2.AbsolutePage = 1
			End If
		End If		 
	end if
	
	if err <> 0 then
		DisConnDBH ObjConn2, ObjRS2	'DB切断
		
		jampErrerP "2","b301","01","実搬出入事前情報","102","SQL:<BR>" & StrSQL & err.description & Err.number
		Exit Function
	end if			

	'エラートラップ解除
    on error goto 0	
End Function

Function LfUpdTruckerAns()
	dim ObjConnUpd, ObjRSUpd, StrSQL
	'エラートラップ開始
	On Error Resume Next


	'DB接続	
	ConnDBH ObjConnUpd, ObjRSUpd
	
	if Num>0 then
		For i=1 To CInt(v_DataCnt1)-1
			if Trim(Request.form("chkAnsNo" & i)) <> "0" then			
				StrSQL = "UPDATE BookingAssign SET UpdtTime='"& Now() &"', UpdtPgCd='TOP', " & _
					   "UpdtTmnl='"& USER &"', TruckerFlag='" & Trim(Request.form("chkAnsNo" & i)) & "' " & _
					   "WHERE BookNo='"& Trim(Request.form("BookNoTbl1" & i)) &"' " & _
					   "AND SenderCode='"& Trim(Request.form("SenderCodeTbl1" & i)) &"' " & _
					   "AND TruckerCode='"& Trim(Request.form("TruckerCodeTbl1" & i)) &"' AND Process='R' "
			
				ObjConnUpd.Execute(StrSQL)
				
				if err <> 0 then
					Set ObjRSUpd = Nothing
					jampErrerPDB ObjConnUpd,"2","b107","01","実搬出:紹介済処理","104","SQL:<BR>"&strSQL
				end if				
			end if
		Next
	end if

	if Num2>0 then		
		For i=1 To CInt(v_DataCnt2)-1						
		if Trim(Request.form("chkAns2No" & i)) <> "0" then		  	
			if CInt(Trim(Request.form("TruckerSubCodeTbl2" & i))) > 0 then
				'2010/04/25 Upd-S Tanaka 搬出搬入によってWkTypeを変更して更新するように修正
				'StrSQL = "UPDATE hITReference SET TruckerFlag" & Trim(Request.form("TruckerSubCodeTbl2" & i)) & "='" & Trim(Request.form("chkAns2No" & i)) & "'," & _
				'		 " UpdtTime='"& Now() &"', UpdtPgCd='TOP',UpdtTmnl='" & USER & "'" & _
				'		 " WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo " & _
				'		 " WHERE WkNo='"& Trim(Request.form("WkNo" & i)) &"' AND WkType='1' AND Process='R')"

				StrSQL = "UPDATE hITReference SET TruckerFlag" & Trim(Request.form("TruckerSubCodeTbl2" & i)) & "='" & Trim(Request.form("chkAns2No" & i)) & "'," & _
						 " UpdtTime='"& Now() &"', UpdtPgCd='TOP',UpdtTmnl='" & USER & "'" & _
						 " WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo " & _
						 " WHERE WkNo='"& Trim(Request.form("WkNo" & i)) & "'"
				'搬出の場合
				If Trim(Request.form("TypeCode" & i))="FULLOUT" Then
					StrSQL =StrSQL & " AND WkType='1' "
				Else
					StrSQL =StrSQL & " AND WkType='3' "
				End If
				StrSQL =StrSQL & " AND Process='R') "
				'2010/04/25 Upd-E Tanaka
				ObjConnUpd.Execute(StrSQL)
				
				if err <> 0 then
					Set ObjRSUpd = Nothing				
					jampErrerPDB ObjConnUpd,"2","b107","01","実搬出:紹介済処理","104","SQL:<BR>"&strSQL
				end if		
			end if		
		end if
		Next	
	end if	
	
	'DB接続解除
	DisConnDBH ObjConnUpd, ObjRSUpd

	'エラートラップ解除
	on error goto 0		  
End Function

'2010/04/27-3 Upd-S C.Pestano
Function getSort1(Key,SortKey,str)
	getSort1 = str
	'2010/04/27 Upd-S C.Pestano
		'if Key <> "" then	
	'		if str = "" then
	'			'getSort1 = " ORDER BY (Case When LTRIM(ISNULL(" & FieldName1(Key) & ",''))='' Then 1 Else 0 End), " & FieldName1(Key) & " " & SortKey
	'			getSort1 = " ORDER BY " & FieldName1(Key) & " " & SortKey
	'		else
	'			'getSort1 = str & " , (Case When LTRIM(ISNULL(" & FieldName1(Key) & ",''))='' Then 1 Else 0 End), " & FieldName1(Key) & " " & SortKey
	'			getSort1 = str & " , " & FieldName1(Key) & " " & SortKey
	'		end if	
	'	end if	
	if str = "" AND Key<>"" then
		str = " ORDER BY "
	elseif str <> "" AND Key<>"" Then 
		str = str & ","		
	end if
	
	if Key <> "" then 
		if FieldName1(CInt(Key)) = "InputDate" AND SortKey = "ASC" then 
			str = str & " MAXDATE ASC,ISNULL(" & FieldName1(CInt(Key)) & ",DATEADD(Year,100,getdate())) " & SortKey	
		elseif FieldName1(CInt(Key)) = "InputDate" AND SortKey = "DESC" then 
			str = str & " MAXDATE DESC,ISNULL(" & FieldName1(CInt(Key)) & ",DATEADD(Year,100,getdate())) " & SortKey	
		else
			str = str & FieldName1(CInt(Key)) & " " & SortKey	
		end if			
	end if
  	getSort1 = str  
end function

Function getSort2(Key,SortKey,str)
	'2010/04/27-3 Upd-S C.Pestano
	getSort2 = str	
	if str = "" AND Key<>"" then
		str = " ORDER BY "
	elseif str <> "" AND Key<>"" Then 
		str = str & ","		
	end if
	'2010/04/28-1 Upd-S C.Pestano
	if Key <> "" then 
			if FieldName2(CInt(Key)) = "DelPermitDate" AND SortKey = "ASC" then 
		        str = str & " DelPermitDate_Sort ASC "
			elseif FieldName2(CInt(Key)) = "CYCut" AND SortKey = "ASC" then 
		        str = str & " CYCut_Sort ASC "
			elseif FieldName2(CInt(Key)) = "FreeTime" AND SortKey = "ASC" then 
				str = str & " FreeTime_Sort ASC "
			elseif FieldName2(CInt(Key)) = "WorkDate" AND SortKey = "ASC" then 
		        str = str & " WorkDate_Sort ASC "
		    else
		        str = str & FieldName2(Key) & " " & SortKey	
		    end if			
	end if	
	'2010/04/28-1 Upd-E C.Pestano
    getSort2 = str  
	'2010/04/27-3 Upd-S C.Pestano
end function
'2010/04/27-3 Upd-E C.Pestano

Function getImage(SortKey)
	getImage = ""
	if SortKey = "ASC" then
		getImage = CONST_ASC	
	else
		getImage = CONST_DESC
	end if	
end function

function LfgetSCAC(keyCode1,arr())
	dim ObjConnSCAC, ObjRSSCAC, StrSQL
    dim cnt

    cnt = 0         '初期化
    LfgetSCAC = ""
	
	'エラートラップ開始
	on error resume next	
	'DB接続	
	ConnDBH ObjConnSCAC, ObjRSSCAC
	
    redim arr(0, 0)
    if trim(keyCode1) <> "" then
        StrSQL = "select Distinct ShipLine from mShipLine where ShipLine = '" & Trim(keyCode1) & "'"
    else
        StrSQL = "select Distinct ShipLine from mShipLine"
    end if
	
    ObjRSSCAC.Open StrSQL, ObjConnSCAC
	
    if not ObjRSSCAC.eof then
        redim arr(ObjRSSCAC.recordcount,0)
        while not ObjRSSCAC.eof
            cnt=cnt+1
            arr(cnt,0)=Trim(ucase(ObjRSSCAC("ShipLine")))
            ObjRSSCAC.movenext
        wend
    end if
	
	DisConnDBH ObjConnSCAC, ObjRSSCAC	
end function

function LfPutPage(rec,page,pagecount,link)
	dim pg, i, j
	dim FirstPage, LastPage	
	'2010/04/26 Add-S Tanaka
	dim PageIndex
	dim PageWkNo
	dim intNextFlag
	dim strParam
	PageIndex=0
	PageWkNo=0
	'2010/04/26 Add-S Tanaka	
	if rec > 0 then	
		'--- カレントページ（補正）
		if pagecount<page then
			page=pagecount
		end if
		
		'2010/04/26 Add-S Tanaka
		'ページIndexを設定
		PageIndex=Fix(page/gcPage)
		if page mod gcPage=0 then
			PageIndex=PageIndex-1
		End If
		
		PageWkNo=((gcPage*PageIndex)+1)-gcPage
		'先頭ページが0より小さい場合は1を設定
		if PageWkNo<=0 Then
			PageWkNo=0
		End If
		
		'パラメータ設定
		If link="pagenum" Then
			strParam="&SCACCode=" & v_SCACCode
		Else
			strParam="&InOutF=" & v_InOutFlag
		End If

		'2010/04/26 Add-E Tanaka

		'--- 総件数、総ページ数 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & strParam & """>最初へ</a>"
			response.write "| &nbsp;"
			'2010/04/26 Upd-S Tanaka
			'response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & page-1 & """>前へ</a>"
            'Y.TAKAKUWA Upd-S 2015-03-12			
			'if PageWkNo<>0 Then
			'	response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """>前へ</a>"
			'Else
			'	response.write "<font style='color:#FFFFFF;'>前へ</font>"
			'End If
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & page - 1  & strParam & """>前へ</a>"
			'Y.TAKAKUWA Upd-E 2015-03-12
			'2010/04/26 Upd-E Tanaka

		else
			response.write "<font style='color:#FFFFFF;'>最初へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>前へ</font>"
		end if        

		'2010/04/26 Upd-S Tanaka
		'--- インデックス
		'if pagecount>1 then
		'	j=page*(-1)+1			
		'	if pagecount<page+gfclng(gcPage / 2) then
		'		if j<(-gfclng(gcPage / 2) - (page+gfclng(gcPage / 2)-pagecount)) then
		'			j=(-gfclng(gcPage / 2) - (page+gfclng(gcPage / 2)-pagecount))
		'		end if
		'	else
		'		if j<-gfclng(gcPage / 2) then
		'			j=-gfclng(gcPage / 2)
		'		end if
		'	end if
		'		 
		'	response.write "&nbsp;| "
		'
		'			for i=1 to gcPage
		'				if pagecount<page+j then
		'			exit for
		'		end if
		'		if j=0 then
		'			response.write "&nbsp;" & (page+j)
		'		else
		'			'2010/04/26 Upd-S Tanaka
		'			'response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?pagenum=" & (page+j) & """>&nbsp;" & (page+j) & "</a>"					
		'			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & (page+j) & """>&nbsp;" & (page+j) & "</a>"					
		'			'2010/04/26 Upd-E Tanaka
		'		end if
		'		j=j+1
		'	next
		'	response.write "| &nbsp;"
		'end if	
		'--- インデックス
		'ページが1ページ以上存在する場合
		if pagecount>1 then
			response.write "| &nbsp;"

			'指定ページ数分ループ
			for i=1 to gcPage
				'ページ数算出
				PageWkNo=(gcPage*PageIndex)+i

				'ページが全ページより大きい場合は処理中断
				if pagecount< PageWkNo then
					PageWkNo=PageWkNo-1
					exit for
				end if
				'現在選択されているページの場合
				if PageWkNo=page then
					response.write "&nbsp;" & PageWkNo
				else
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """ >&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If
		'2010/04/26 Upd-E Tanaka
		'Y.TAKAKUWA Debug-S 2015-03-12
        'response.write "PAGE: " + CSTR(page + " PAGECOUNT:" + CSTR(pagecount - 1))
        'Y.TAKAKUWA Debug-E 2015-03-12
        'Y.TAKAKUWA Upd-S 2015-03-12
		'if page<pagecount-1 then
		if page<pagecount then
		'Y.TAKAKUWA Upd-E 2015-03-12
			'2010/04/26 Upd-S Tanaka
			'response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & page+1 & """>次へ</a>"'
			'Y.TAKAKUWA Upd-S 2015-03-13
			'PageWkNo=PageWkNo+1
			PageWkNo=page+1
			'Y.TAKAKUWA Upd-E 2015-03-13
			'Y.TAKAKUWA Upd-S 2015-03-12
			'If PageWkNo<=LastPage Then
			'	response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """>次へ</a>"'
			'Else
			'	response.write "<font style='color:#FFFFFF;'>次へ</font>"
			'End If
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & strParam & """>次へ</a>"'
			'Y.TAKAKUWA Upd-E 2015-03-12
			'2010/04/26 Upd-E Tanaka
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & strParam & """>最後へ</a>"'            
		else
			response.write "<font style='color:#FFFFFF;'>次へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>最後へ</font>"
		end if
	end if
end function
'-----------------------------
'   数値変換 (Long型)
'-----------------------------
function gfCLng(str1)
    dim str
    str=gfTrim(str1)
    if isnull(str) then
        gfCLng=0
    elseif trim(str)="" then
        gfCLng=0
    elseif not isNumeric(str) then
        gfCLng=0
    elseif len(str)>9 then
        if instr(str,".")>0 and instr(str,".")<10 then
            gfClng=clng(left(str,instr(str,".")-1))
        else
            gfClng=0
        end if
    else
        gfCLng = CLng(fix(str))
    end if
end function
'-----------------------------
'   Trim　NULLの場合→空値(Space0)
'-----------------------------
function gfTrim(str)
    if isnull(str) then
        gfTrim=""
    else
        gfTrim=trim(str)
    end if
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR>
<STYLE>
	th.hlist {
		position: relative;
	}
	th {
		border-width: 1px 1px 1px 1px;
		padding: 4px;
		background-color: #ffcc33;
	}
	SELECT.chr {
		BACKGROUND-COLOR: #ffffff;
		BORDER-BOTTOM: #ffffff 1px solid;
		BORDER-LEFT: #002f7b 0px solid;
		BORDER-RIGHT: #ffffff 0px solid;
		BORDER-TOP: #ffffff 0px solid;
		COLOR: black;
		FONT-FAMILY: 'ＭＳ ゴシック';
		FONT-SIZE: 12px;
		FONT-WEIGHT: normal;
		PADDING-BOTTOM: 2px;
		PADDING-LEFT: 1px;
		PADDING-RIGHT: 2px;
		PADDING-TOP: 3px;
		TEXT-ALIGN: left
	}
	table {
		border-width: 0px 1px 1px 0px;
	}
	DIV.center {
		text-align:center;
	}
	DIV.BDIV1 {
		position: relative;
		border-width: 0px 0px 1px 0px;
	}
	/*Y.TAKAKUWA Add-S 2015-03-10*/
	DIV.BDIV1HEADER {
		position: relative;
		border-width: 0px 0px 1px 0px;
	}
	/*Y.TAKAKUWA Add-E 2015-03-10*/
	DIV.BDIV2 {
		position: relative
		border-width: 0px 0px 1px 0px;
	}
	/*Y.TAKAKUWA Add-S 2015-03-10*/
	DIV.BDIV2HEADER {
		position: relative;
		border-width: 0px 0px 1px 0px;
	}
	/*Y.TAKAKUWA Add-E 2015-03-10*/
	/*DIV.BDIV3 {
		position: absolute
		border-width: 0px 0px 1px 0px;
	}*/
	thead tr {
		position: relative;
		top: expression(this.offsetParent.scrollTop);
	}
	#loading2 {
		font:bold 12px Verdana;
		color:red;
		position:absolute; 
		top:220px; 
		left:390px;
		width:300px;
		height:30px; 
		z-index:69;
		font-size:12pt;
		border:0px;
		vertical-align: middle;
	}
	.th-inner {
		position: absolute;
		top: 0;
		line-height: 30px; /* height of header */
		text-align: left;
		border-left: 1px solid black;
		padding-left: 5px;
		margin-left: -5px;
	}

</style>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT Language="JavaScript">
	/*
	function finit(){
		var str;
		
		//データ引継ぎ設定  
		document.frm.Gamen_Mode.value="<%=v_GamenMode%>";    //処理モード
		document.frm.SCACSrhFlag.value='';	//2010/4/26 Add Tanaka
		document.frm.InOutSrhFlag.value='';	//2010/4/26 Add Tanaka
		str = readCookie('HitsTbl1')
		
		if(str!= null && "<%=Num%>" != "0"){  		  		
			displayColumn(str,0,"H1Col01","D1Col01")
			displayColumn(str,1,"H1Col02","D1Col02")
			displayColumn(str,2,"H1Col03","D1Col03")
			displayColumn(str,3,"H1Col04","D1Col04")
			displayColumn(str,4,"H1Col05","D1Col05")
			displayColumn(str,5,"H1Col06","D1Col06")
			displayColumn(str,6,"H1Col07","D1Col07")
			displayColumn(str,7,"H1Col08","D1Col08")
			displayColumn(str,8,"H1Col09","D1Col09")
			displayColumn(str,9,"H1Col10","D1Col10")
			displayColumn(str,10,"H1Col11","D1Col11")
			displayColumn(str,11,"H1Col12","D1Col12")
			displayColumn(str,12,"H1Col13","D1Col13")
			displayColumn(str,13,"H1Col14","D1Col14")
		}
		
		str = readCookie('HitsTbl2')
		
		if(str!= null && "<%=Num2%>" != "0"){  		  		
			displayColumn(str,0,"H2Col01","D2Col01")
			displayColumn(str,1,"H2Col02","D2Col02")
			displayColumn(str,2,"H2Col03","D2Col03")
			displayColumn(str,3,"H2Col04","D2Col04")
			displayColumn(str,4,"H2Col05","D2Col05")
			displayColumn(str,5,"H2Col06","D2Col06")
			displayColumn(str,6,"H2Col07","D2Col07")
			displayColumn(str,7,"H2Col08","D2Col08")
			displayColumn(str,8,"H2Col09","D2Col09")
			displayColumn(str,9,"H2Col10","D2Col10")
			displayColumn(str,10,"H2Col11","D2Col11")
			displayColumn(str,11,"H2Col12","D2Col12")
			displayColumn(str,12,"H2Col13","D2Col13")
			displayColumn(str,13,"H2Col14","D2Col14")
			displayColumn(str,14,"H2Col15","D2Col15")
		}
	}

	function displayColumn(str,colNo,colHeader,colData){
		var disp;

		if(str.charAt(colNo) == "0"){
			disp = 'none'
		}else{
			disp = '';
		}
		
		document.getElementById(colHeader).style.display = disp;
		
		var oObject = document.all.tags("td");
		var i;
		for (i=0; i<=oObject.length-1; i++){
			if (oObject[i].id.substring(0,7) == colData) {
				oObject[i].style.display=disp;				
			}
		}
	}
	*/

	function finit(){
		var str;
		
		//データ引継ぎ設定  
		document.frm.Gamen_Mode.value="<%=v_GamenMode%>";    //処理モード
		document.frm.SCACSrhFlag.value='';	//2010/4/26 Add Tanaka
		document.frm.InOutSrhFlag.value='';	//2010/4/26 Add Tanaka
		str = readCookie('HitsTbl1')
		
		if(str!= null && "<%=Num%>" != "0"){  		  		
			displayColumn(str,"TBEmpty")
		}
		
		str = readCookie('HitsTbl2')
		
		if(str!= null && "<%=Num2%>" != "0"){  		  		
			displayColumn(str,"TBInOut")
		}
	}

	function displayColumn(str,TableName){
		var objTable;
		var intMaxRow;
		var disp;
		var colHeader;

		//Tableオブジェクト取得
		objTable=document.getElementById(TableName);
		//行数取得
		intMaxRow=objTable.rows.length;

		var trs=objTable.getElementsByTagName("TR");


		if (TableName=='TBEmpty'){
			//行数分ループ
			for(var intRowCnt=1; intRowCnt<intMaxRow; intRowCnt++) {

				var tds=trs[intRowCnt].getElementsByTagName("TD");
				for(var intColCnt=0; intColCnt<14; intColCnt++) {

					//ヘッダーIDを設定
					if(intColCnt<9){
						colHeader="H1Col0"+(intColCnt+1)
					}else{
						colHeader="H1Col"+(intColCnt+1)
					}


					if(str.charAt(intColCnt) == "0"){
						disp = 'none'
					}else{
						disp = '';
					}
					if (intRowCnt==1){
						document.getElementById(colHeader).style.display = disp;
					}
					tds[intColCnt].style.display =disp;
				}

			}

		}else{
			//行数分ループ
			for(var intRowCnt=1; intRowCnt<intMaxRow; intRowCnt++) {

				var tds=trs[intRowCnt].getElementsByTagName("TD");
				//Y.TAKAKUWA Upd-S 2013-02-18
				for(var intColCnt=0; intColCnt<19; intColCnt++) { //17->19
					//ヘッダーIDを設定
					if(intColCnt<9){
						colHeader="H2Col0"+(intColCnt+1)
					}else{
						colHeader="H2Col"+(intColCnt+1)
					}
					if(str.charAt(intColCnt) == "0"){
						disp = 'none'
					}else{
						disp = '';
					}
					if (intRowCnt==1){
						document.getElementById(colHeader).style.display = disp;
					}
					tds[intColCnt].style.display =disp;
				}
				//Y.TAKAKUWA Upd-E 2013-02-18

			}	

		}

	}


	function readCookie(name) {
		var nameEQ = name + "=";
		var ca = document.cookie.split(';');
		for(var i=0;i < ca.length;i++) {
			var c = ca[i];
			while (c.charAt(0)==' ') c = c.substring(1,c.length);
			if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
		}
		return null;
	}

	//JC.MARINDUQUE Add-S 03-14-2014
	function getInternetExplorerVersion()
	{
		var rv = -1;
		if (navigator.appName == 'Microsoft Internet Explorer')
		{
			var ua = navigator.userAgent;
			var re  = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
			if (re.exec(ua) != null)
			rv = parseFloat( RegExp.$1 );
		}
		else if (navigator.appName == 'Netscape')
		{
			var ua = navigator.userAgent;
			var re  = new RegExp("Trident/.*rv:([0-9]{1,}[\.0-9]{0,})");
			if (re.exec(ua) != null)
			rv = parseFloat( RegExp.$1 );
		}
		
		return rv;
	}

	//JC.MARINDUQUE Add-E 03-14-2014

	//データが無い場合の表示制御
	function view(){
		//JC.MARINDUQUE Add-S 03-14-2014
		var IEVersion = getInternetExplorerVersion();
		//JC.MARINDUQUE Add-E 03-14-2014
		
		//JC.MARINDUQUE Edit-S 03-14-2014
		var obj3=document.getElementById("BDIV3");

		if((document.body.offsetWidth-10)  < 880){
			obj3.style.width=document.body.offsetWidth-10;
			obj3.style.overflowX="auto";
		}
		else{
			obj3.style.width=document.body.offsetWidth-10;
			obj3.style.overflowX="auto";
		}

		obj3.style.height=document.body.offsetHeight;
		obj3.style.overflowY="auto";
		
		var obj1=document.getElementById("BDIV1");
		
		if(IEVersion < 10)
		{
			if((document.body.offsetWidth-200) < 50){
				obj1.style.width=50;
				obj1.style.overflowX="auto";	 
			}else if((document.body.offsetWidth-200)  < 813){
				//obj1.style.width=document.body.offsetWidth-200;
				obj1.style.width=document.body.offsetWidth-30;
				obj1.style.overflowX="auto";
			}else{
				//obj1.style.width=document.body.offsetWidth-200;
				obj1.style.width=document.body.offsetWidth-30;
				obj1.style.overflowX="auto";
			}
			
			if(document.body.offsetHeight-100 < 50){
				obj1.style.height = 50;
				obj1.style.overflowY = "auto";
			}else if(document.body.offsetHeight-100 < 150){
				obj1.style.height = document.body.offsetHeight-115;
				obj1.style.overflowY = "auto";	
			}else{
				obj1.style.height = 200; // Edited by AK.DELAROSA 2021/01/12
				obj1.style.overflowY = "auto";
			}		
		}
		else
		{
			//JC.MARINDUQUE Add-S 03-14-2014
			var initialHeight = document.documentElement.clientHeight;
			
			if((document.body.offsetWidth-200) < 50){
				obj1.style.width=50;
				obj1.style.overflowX="auto";	 
			}else if((document.body.offsetWidth-200)  < 813){
				//obj1.style.width=document.body.offsetWidth-200;
				obj1.style.width=document.body.offsetWidth-30;
				obj1.style.overflowX="auto";
			}else{
				//obj1.style.width=document.body.offsetWidth-200;
				obj1.style.width=document.body.offsetWidth-30;
				obj1.style.overflowX="auto";
			}
			
			if(initialHeight-100 < 50){
				obj1.style.height = 50;
				obj1.style.overflowY = "auto";
			}else if(initialHeight-100 < 150){
				obj1.style.height = initialHeight-115;
				obj1.style.overflowY = "auto";	
			}else{
				obj1.style.height = 150;
				obj1.style.overflowY = "auto";
			}
			//JC.MARINDUQUE Add-E 03-14-2014
		}	

		//Y.TAKAKUWA ADD-S 2015-03-06
		var obj1header=document.getElementById("BDIV1HEADER");
		if(IEVersion < 10)
		{
			if((document.body.offsetWidth-200) < 50){
				obj1header.style.width=50; 
			}else if((document.body.offsetWidth-200)  < 813){
				if(IEVersion == 7){
					obj1header.style.width=obj1.clientWidth;//document.body.offsetWidth-45;
					obj1header.style.position = "relative";
				}
				else{
					obj1header.style.width=obj1.clientWidth;//document.body.offsetWidth-45;
				}
			}else{
				if(IEVersion == 7){
					obj1header.style.width=obj1.clientWidth;//document.body.offsetWidth-45;
					obj1header.style.position = "relative";
				}
				else{
					obj1header.style.width=obj1.clientWidth;//document.body.offsetWidth-45;
				}
			}	
			obj1header.style.height = 35;
		}
		else
		{
			
			if((document.body.offsetWidth-200) < 50){
				obj1header.style.width=50;	 
			}else if((document.body.offsetWidth-200)  < 813){
				obj1header.style.width=obj1.clientWidth;//document.body.offsetWidth-45;
			}else{
				obj1header.style.width=obj1.clientWidth;//document.body.offsetWidth-45;
			}
			obj1header.style.height = 35;
		}	
		//Y.TAKAKUWA ADD-E 2015-03-06
		
		var obj2=document.getElementById("BDIV2");
		
		if(IEVersion < 10)
		{
			if((document.body.offsetWidth-200) < 50){
				obj2.style.width=50;
				obj2.style.overflowX="auto";	 
			}else if((document.body.offsetWidth-200)  < 813){
				obj2.style.width=document.body.offsetWidth-30;
				obj2.style.overflowX="auto";
			}else{
				obj2.style.width=document.body.offsetWidth-30;
				obj2.style.overflowX="auto";
			}	
			
			if((document.body.offsetHeight-280) < 50){
				obj2.style.height = 50;
				obj2.styleoverflowY = "auto";	 
			}else if((document.body.offsetHeight-280)  < 335){
				obj2.style.height = document.body.offsetHeight-410; //350 Edited by AK.DELAROSA 2021/01/12
				obj2.style.overflowY = "auto";
			}else if((document.body.offsetHeight-310)  < 395){
				obj2.style.height = document.body.offsetHeight-450;
				obj2.style.overflowY = "auto";
			}else{
				obj2.style.height = 310; // Edited by AK.DELAROSA 2021/01/12
				obj2.style.overflowY = "auto";
			}
		}
		else
		{
		//JC.MARINDUQUE Add-S 03-14-2014
			var initialHeight = document.documentElement.clientHeight;
			
			if((document.body.offsetWidth-200) < 50){
				obj2.style.width=50;
				obj2.style.overflowX="auto";	 
			}else if((document.body.offsetWidth-200)  < 813){
				obj2.style.width=document.body.offsetWidth-200;
				obj2.style.width=document.body.offsetWidth-30;
				obj2.style.overflowX="auto";
			}else{
				obj2.style.width=document.body.offsetWidth-200;
				obj2.style.width=document.body.offsetWidth-30;
				obj2.style.overflowX="auto";
			}

			if((initialHeight-280) < 50){
				obj2.style.height = 50;
				obj2.styleoverflowY = "auto";	 
			}else if((initialHeight-280)  < 335){
				//Y.TAKAKUWA Upd-S 2015-03-06
				//obj2.style.height = initialHeight-280;
				obj2.style.height = initialHeight-350;
				//Y.TAKAKUWA Upd-E 2015-03-06
				obj2.style.overflowY = "auto";
			}else{
				//Y.TAKAKUWA Upd-S 2015-03-06
				//obj2.style.height = initialHeight-280;
				obj2.style.height = initialHeight-350;
				//Y.TAKAKUWA Upd-E 2015-03-06
				obj2.style.overflowY = "auto";
			}
			
		//JC.MARINDUQUE Add-E 03-14-2014
		}
		//JC.MARINDUQUE Edit-E 03-14-2014
		
		//Y.TAKAKUWA Add-S 2015-03-06
		var obj2header=document.getElementById("BDIV2HEADER");
		if(IEVersion < 10)
		{
			if((document.body.offsetWidth-200) < 50){
				obj2header.style.width=50; 
			}else if((document.body.offsetWidth-200)  < 813){
				obj2header.style.width=obj2.clientWidth;//document.body.offsetWidth-45;
				if(IEVersion == 7){
					//obj2header.style.position = "relative";
				}
				else{
					obj2header.style.width=obj2.clientWidth;//document.body.offsetWidth-45;
				}
			}else{
				if(IEVersion == 7){
					obj2header.style.width=obj2.clientWidth;//document.body.offsetWidth-45;
					//obj2header.style.position = "relative";
				}
				else{
					obj2header.style.width=obj2.clientWidth;//document.body.offsetWidth-45;
				}
			}	
			obj2header.style.height = 35;
		}
		else
		{
			var initialHeight = document.documentElement.clientHeight;
			
			if((document.body.offsetWidth-200) < 50){
				obj2header.style.width=50;
			}else if((document.body.offsetWidth-200)  < 813){
				obj2header.style.width=obj2.clientWidth;//document.body.offsetWidth-45;
			}else{
				obj2header.style.width=obj2.clientWidth;//document.body.offsetWidth-45;
			}
			obj2header.style.height = 35;
		}
		
		//Y.TAKAKUWA Add-S 2015-03-06

	}

	function OpenCodeWin(strTitle,tableno)
	{
		var CodeWin;
		var w=400;
		var h=525;
		var l=0;
		var t=0;
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
		
		CodeWin = window.open("./show.asp?user=<%=Session.Contents("userid")%>&pagetitle="+strTitle+"&show_column="+tableno,"codelist","scrollbars=yes,resizable=yes,width="+w+",height="+h+",top="+t+",left="+l);

		CodeWin.focus();

	}

	function OpenCodeWin2(tableno){
		var CodeWin;
		var w=400;
		var h=300;
		var l=0;
		var t=0;
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
		
		CodeWin = window.open("./sort.asp?user=<%=Session.Contents("userid")%>&left_menu="+tableno,"codelist","scrollbars=yes,resizable=yes,width="+w+",height="+h+",top="+t+",left="+l);
		CodeWin.focus();
	}

	//検索
	function SerchC1(SortFlag,Key){
		Fname=document.frm;
		Fname.SortFlag1.value=SortFlag;
		Fname.SortKey1.value=Key;
		Fname.Gamen_Mode.value = "T1"
		Fname.submit();
	}

	//検索
	function SerchC2(SortFlag,Key){
		Fname=document.frm;
		Fname.SortFlag2.value=SortFlag;
		Fname.SortKey2.value=Key;
		Fname.Gamen_Mode.value = "T2"
		Fname.submit();
	}

	function openwin(){

		var w=1400;
		var h=750;
		var l=0;
		var t=0;
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
		
		var win=window.open("../download/download_list.asp","","status=no,width="+w+",height="+h+",top="+t+",left="+l);
	}

	function UpdAns(){
		document.frm.Gamen_Mode.value = "U";
		document.frm.submit();
	}

	function fPrint(){
		var cnt;
		cnt = 0;
		for (i = 1; i <= "<%=ObjRS2.PageSize%>"; i++) {
			if (chkobj("chkInOut" + i)) {  //チェックボックスがチェックされている場合
				cnt++;
			}
		}
		
		if(cnt == 0) {
			window.alert("搬入票出力対象チェックが選択されていません。");
			return false;
		}
		
		document.frm.Gamen_Mode.value = "P";
		document.frm.submit();
	}

	//引数ｉｄがチェックされているかどうかを確認
	//戻り値：１ チェックされている
	//　　　　０ チェックされていない
	function chkobj(id)
	{
		var obj;
		obj = eval("document.frm." + id);
		if(obj != null){
			return (obj.checked) ? 1 : 0;
		}
	}

	//更新
	function GoRenew(No){
		Fname=document.frm;
		Fname.targetNo.value=No;
		Fname.action="./dmi015.asp";
		newWin = window.open("", "ReEntry", "width=900,height=500,left=10,top=10,status=yes,resizable=yes,scrollbars=yes");
		Fname.target="ReEntry";
		Fname.elements[No].disabled=false;
		Fname.submit();
		Fname.elements[No].disabled=true;
		Fname.target="_self";
		//2010/04/25 Add-S Tanaka
		Fname.action="";
		//2010/04/25 Add-E Tanaka
	}
	function GoRenewEmpty(bookNo,compF,SijiM,SijiC,sShipLine){
	Fname=document.frm;
	Fname.BookNo.value=bookNo;
	Fname.CompF.value=compF;
	Fname.COMPcd0.value=SijiM;
	Fname.COMPcd1.value=SijiC;
	Fname.ShipLine.value=sShipLine;
	Fname.action="./dmi312.asp";
	newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes,scrollbars=yes");
	Fname.target="ReEntry";
	Fname.submit();  
	//2010/04/25 Add-S Tanaka
	Fname.target="_self";
	Fname.action="";
	//2010/04/25 Add-E Tanaka
	}
	//更新
	function GoRenew2(sakuNo,bookNo,conNo){
	Fname=document.frm;
	Fname.SakuNo.value=sakuNo;
	Fname.BookNo.value=bookNo;
	Fname.CONnum.value=conNo;
	Fname.action="./dmo320.asp";
	newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes,scrollbars=yes");
	Fname.target="ReEntry";
	Fname.submit();
	//2010/04/25 Add-S Tanaka
	Fname.target="_self";
	Fname.action="";
	//2010/04/25 Add-E Tanaka
	}

	//コンテナ詳細
	function GoConinf2(wkcNo,flag,conNo){
		Fname=document.frm;  
		Fname.SakuNo.value=wkcNo;
		Fname.flag.value=flag;
		Fname.CONnum.value=conNo;
		Fname.action="./dmo900.asp";
		newWin = window.open("", "ConInfo", "left=30,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes");
		Fname.target="ConInfo";
		Fname.submit();
		Fname.target="_self";
		Fname.InfoFlag.value="0";
		//2010/04/25 Add-S Tanaka
		Fname.action="";
		//2010/04/25 Add-E Tanaka
	}

	//コンテナ一覧
	function GoConinf(BookNo){
		var w=300;
		var h=500;
		var l=0;
		var t=0;
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
		
		var win=window.open("./contlist.asp?BookNo="+BookNo,"","status=no,width="+w+",height="+h+",top="+t+",left="+l);
		
	}
	//コンテナ詳細
	function GoConinfContNo(conNo){
	Fname=document.frm;
	Fname.CONnum.value=conNo;
	Fname.BookNo.value="";        //CW-021 ADD
	BookInfo(Fname);
	//2010/04/25 Add-S Tanaka
	Fname.action="";
	//2010/04/25 Add-E Tanaka
	}
	function showContent(){
		var target=null;
		
		while (target==null) {
			target=parent.window.frames(0);
		}
		
		var target1 = document.getElementById("loading2");
		target1.style.display='none';
		
		//show content
		document.getElementById("content").style.display='block';
	}

	//2010/04/25 Add-S Tanaka
	function fSCACCode(){
		document.frm.Gamen_Mode.value = "";
		document.frm.SCACSrhFlag.value="1";
		document.frm.submit();
	}
	function fInOut(){
		document.frm.Gamen_Mode.value = "";
		document.frm.InOutSrhFlag.value="1";
		document.frm.submit();
	}
	function flReload(){
		parent.Top.location.reload(true);
		document.frm.Gamen_Mode.value = "R";
		document.frm.submit();
	}
	//2010/04/25 Add-E Tanaka
	//Y.TAKAKUWA Add-S 2015-03-06
	function cloneTable(tblSource, tblDestination, type)
	{
		if(type=="1")
		{
			<%If Num<>0 Then%> 
			//Y.TAKAKUWA Upd-S 2015-04-06
			//Change the name of cloned elements
			var source = document.getElementById(tblSource);
			var destination = document.getElementById(tblDestination);
			var copy = source.cloneNode(true);
			var rowCount = copy.rows.length;
			for(var i=1; i<rowCount; i++) {
				var row = copy.rows[i];
				element_i = row.cells[3].getElementsByTagName('select')[0];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[0];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[1];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[2];
				element_i.removeAttribute('name');
			}
			destination.parentNode.replaceChild(copy, destination);
			source.style.marginTop = "-35px";
			//Y.TAKAKUWA Upd-E 2015-04-06
			<%end if%>
		}
		else
		{
			<%If Num2<>0 Then%> 
			//Y.TAKAKUWA Upd-S 2015-04-06
			//Change the name of cloned elements
			var source = document.getElementById(tblSource);
			var destination = document.getElementById(tblDestination);
			var copy = source.cloneNode(true);
			var rowCount = copy.rows.length;
			for(var i=1; i<rowCount; i++) {
				var row = copy.rows[i];
				element_i = row.cells[0].getElementsByTagName ('input')[0];
				element_i.removeAttribute('name');
				element_i = row.cells[6].getElementsByTagName('select')[0];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[2];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[3];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[4];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[5];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[6];
				element_i.removeAttribute('name');
				element_i = row.getElementsByTagName ('input')[7];
				element_i.removeAttribute('name');
			}
			destination.parentNode.replaceChild(copy, destination);
			source.style.marginTop = "-35px";
			//Y.TAKAKUWA Upd-E 2015-04-06
			<%end if%>
		}
	}
	function onScrollDiv(Scrollablediv,Scrolleddiv) {
		document.getElementById(Scrolleddiv).scrollLeft = Scrollablediv.scrollLeft;
	}
	//Y.TAKAKUWA Add-E 2015-03-06
</SCRIPT>
</HEAD>
<!-- 2010/04/27-2 Upd-S C.Pestano-->
<BODY onLoad="finit();view();" onResize="view();">
	<!--BODY onLoad="setTimeout('showContent()', 500);finit();view();" onResize="view();"-->
	<!--div class="center" id="loading2">しばらくお待ちください。&nbsp;<IMG border=0 src=Image/loaded.gif></div-->
	<!--div id="content" style="display:none;"-->
	<!-- 2010/04/27-2 Upd-E C.Pestano-->
	<form name="frm" method="post">
		<div id="BDIV3">
			<INPUT type=hidden name="SortFlag1">
			<INPUT type=hidden name="SortFlag2">
			<INPUT type=hidden name="SortKey1" value="<%=v_SortKey1%>">
			<INPUT type=hidden name="SortKey2" value="<%=v_SortKey2%>">
			<INPUT type=hidden name="Gamen_Mode" size="9" readonly tabindex= -1>
			<INPUT type=hidden name="InfoFlag" value="">
			<INPUT type=hidden name="SakuNo" value="">
			<INPUT type=hidden name="flag" value="">
			<INPUT type=hidden name="targetNo" value="">
			<INPUT type=hidden name="CONnum" value="" >
			<INPUT type=hidden name="BookNo" value="" >
			<INPUT type=hidden name="CompF" value="" >
			<INPUT type=hidden name="COMPcd0" value="" >
			<INPUT type=hidden name="COMPcd1" value="" >
			<INPUT type=hidden name="ShipLine" value="" >
			<INPUT type=hidden name="ShoriMode" value="EMoutUpd">
			<INPUT type=hidden name="Mord" value="1" >
			<INPUT type=hidden name="SCACSrhFlag">
			<INPUT type=hidden name="InOutSrhFlag">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>		
					<td width="100%">
						<table cellpadding="1" cellspacing="0">
							<tr>
								<td width="20"></td>
								<td width="125" nowrap><a href="../userchk.asp?link=SendStatus/sst000F.asp" target="_blank">輸入ステータス配信</a></td>
								<td width="125" nowrap><a href="../userchk.asp?link=terminal.asp" target="_blank">CY混雑状況・映像</a></td>
								<td width="125" nowrap><a href="../userchk.asp?link=arvdepinfo.asp" target="_blank">離着岸情報照会</a></td>
								<td width="125" nowrap><a href="JavaScript:openwin()">利用者ガイド</a></td>
								<!--td><input name="btn3" type="button" value="更新" onClick="javascript:parent.List.location.reload(true);parent.Top.location.reload(true);"></td-->
								<td width="350" nowrap></td>
								<td width="125"><input name="btn3" type="button" style="WIDTH: 150px;" value="作業テーブルデータ更新" onClick="flReload();"></td><!-- 2010/05/06 Add-E C.Pestano-->
							</tr>									
							<tr>
								<td width="20"></td>
								<td nowrap><B>空バンピック情報</B></td>
								<td>
									<select name="cmbSCACCode" onChange="fSCACCode();">
									<% 
										v_work = ""                  							
										Response.Write "<OPTION VALUE = ''>　"
										for v_loop = 1 to  ubound(Arr_SCACCode)
											v_work = Arr_SCACCode(v_loop, 0)
											Response.Write "<OPTION VALUE ='" & v_work & "'"
											if v_work = v_SCACCode then Response.Write " SELECTED"
											Response.write ">" & Arr_SCACCode(v_loop, 0)
										next 							
									%>
									</select>
								</td>
								<td><input name="btn1" type="button" value="表示列設定" onClick="OpenCodeWin('表示列選択（空搬出）','1')"></td>
								<td><input name="btn2" type="button" value="並び替え" onClick="OpenCodeWin2('5');"></td>
								<!-- 2010/04/25 Upd-S C.Pestano -->
								<!--td><input name="btn3" type="button" value="更新" onClick="javascript:parent.List.location.reload(true);parent.Top.location.reload(true);"></td-->
								<!-- 2010/04/25 Upd-S C.Pestano -->
								<!--td width="50">&nbsp;</td-->
								<!--Y.TAKAKUWA Upd-S 2015-03-11-->
								<!--<td>-->
								<td width="600">
								<!--Y.TAKAKUWA Upd-E 2015-03-11  -->
								<%														
									
									
									if Num > 0 then
										abspage = ObjRS.AbsolutePage
										pagecnt = ObjRS.PageCount	
										
										call LfPutPage(Num,abspage,pagecnt,"pagenum")					
									end if
									'If Num > 100 Then
										'	abspage = ObjRS.AbsolutePage
										'	pagecnt = ObjRS.PageCount							                     		
										'							
										'	Response.Write "<div align=""center"">" & vbcrlf
										'	Response.Write "<a href="""
										'	Response.Write Request.ServerVariables("SCRIPT_NAME")
										'	Response.Write "?pagenum=1""><b>最初のページ</b></a>"
										'	Response.Write "	|	"								
										'														
										'If abspage = 1 Then
										'		Response.Write "<span>" & abspage & "</span>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 4 & """>&nbsp;<b>" & abspage + 4 & "</b></a>"
										'Elseif abspage = 2 then
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"							
										'		Response.Write "<span>&nbsp;" & abspage & "</span>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
										'Elseif abspage = pagecnt then
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 4 & """>&nbsp;<b>" & abspage - 4 & "</b></a>"															
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"															
										'		Response.Write "<span>&nbsp;" & abspage & "</span>"
										'Elseif abspage = CInt(pagecnt-1) then
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"															
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
										'		Response.Write "<span>&nbsp;" & abspage & "</span>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"									
										'Else
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
										'		Response.Write "<span>&nbsp;" & abspage & "</span>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
										'		Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'		Response.Write "?pagenum=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"								
										'End If
										'							
										'	Response.Write "	|	"
										'	Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	Response.Write "?pagenum=" & pagecnt & """><b>最後のページ</b></a>"
										'	Response.Write "</div>" & vbcrlf						
										'										
									'End If
								%>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>		
					<td>
						<!--Y.TAKAKUWA Add-S 2015-03-05-->	
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<tr>
						<td>
							<DIV ID="BDIV1HEADER" style="overflow:hidden;">
							<table border="1" cellpadding="0" cellspacing="0" width="100%" Id="TBEmpty1">										
							
							</table>
							</DIV>
						</td>
						</tr>
						<tr>
						<td>
						<!--Y.TAKAKUWA Add-E 2015-03-05-->	
						<div id="BDIV1" onscroll="onScrollDiv(this,'BDIV1HEADER');"  style="zoom: 1">						
							<% If Num>0 Then%>
							<% If blnSorted Then%>
							<!--<iframe frameborder="0" style="background-color:transparent;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400px;height:48px;"></iframe>-->	
							<% Else%>
							<!--<iframe frameborder="0" style="background-color:transparent;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400px;height:38px;"></iframe>-->	
							<% End If%>
								
								<table border="1" cellpadding="0" cellspacing="0" width="100%" Id="TBEmpty">
									<thead>																				
										<th id="H1Col01" class="hlist" nowrap style="width:50px;"><%=HeaderTbl1(0)%></th>							
										<th id="H1Col02" class="hlist" nowrap ><%=HeaderTbl1(1)%></th>
										<th id="H1Col03" class="hlist" nowrap ><%=HeaderTbl1(2)%></th>
										<th id="H1Col04" class="hlist" nowrap ><%=HeaderTbl1(3)%></th>	
										<th id="H1Col05" class="hlist" nowrap ><%=HeaderTbl1(4)%></th>							
										<th id="H1Col06" class="hlist" nowrap ><%=HeaderTbl1(5)%></th>
										<th id="H1Col07" class="hlist" nowrap ><%=HeaderTbl1(6)%></th>
										<th id="H1Col08" class="hlist" nowrap ><%=HeaderTbl1(7)%></th>
										<th id="H1Col09" class="hlist" nowrap ><%=HeaderTbl1(8)%></th>
										<th id="H1Col10" class="hlist" nowrap ><%=HeaderTbl1(9)%></th>
										<th id="H1Col11" class="hlist" nowrap ><%=HeaderTbl1(10)%></th>
										<th id="H1Col12" class="hlist" nowrap ><%=HeaderTbl1(11)%></th>
										<th id="H1Col13" class="hlist" nowrap ><%=HeaderTbl1(12)%></th>
										<th id="H1Col14" class="hlist" nowrap ><%=HeaderTbl1(13)%></th>
									</thead>
									<tbody>							
										<% 
											x = 1
											For i=1 To ObjRS.PageSize
												If Not ObjRS.EOF Then
												x = x + 1								

										%>								
										<tr class=bgw>													
											<td id="D1Col01" height="22" nowrap><%=Trim(ObjRS("InPutDate"))%><BR></td>	
											<td id="D1Col02" nowrap ><%=Trim(ObjRS("SenderCode"))%><BR></td>
											<td id="D1Col03" nowrap ><%=Trim(ObjRS("Flag2"))%><BR></td>
											<td id="D1Col04" nowrap >								
											<%
												v_ItemName = "chkAnsNo" + cstr(i)
												Response.Write "<select name= '" & v_ItemName & "' class=chr>"													
												Response.Write "<option value='0'>未</option>"
												Response.Write "<option value='1'>Yes</option>"
												Response.Write "<option value='2'>No</option>"										
												Response.Write "</select>"										
											%>								
											</td>
											<!--td id="D1Col05" nowrap><%=Trim(ObjRS("BookNo"))%><BR></td-->
											<td id="D1Col05" nowrap ><A HREF="JavaScript:GoRenewEmpty('<%=Trim(ObjRS("BookNo"))%>','<%=Trim(ObjRS("NumCount"))%>','<%=Trim(ObjRS("SenderCode1"))%>','<%=Trim(ObjRS("TruckerCode"))%>','<%=Trim(ObjRS("ShipLine"))%>');"><%=Trim(ObjRS("BookNo"))%></A><BR></td>
											<td id="D1Col06" nowrap><A HREF="JavaScript:GoConinf('<%=Trim(ObjRS("BookNo"))%>');"><%=Trim(ObjRS("NumCount"))%></A><BR></td>								
											<!--td id="D1Col06" nowrap><%=Trim(ObjRS("NumCount"))%><BR></td-->								
											<td id="D1Col07" nowrap><%=Trim(ObjRS("ContSize1"))%><BR></td>
											<td id="D1Col08" nowrap><%=Trim(ObjRS("ContType1"))%><BR></td>
											<td id="D1Col09" nowrap><%=Trim(ObjRS("ContHeight1"))%><BR></td>
											<td id="D1Col10" nowrap><%=Trim(ObjRS("ContMaterial1"))%><BR></td>
											<td id="D1Col11" nowrap><%=Trim(ObjRS("ShipLine"))%><BR></td>
											<td id="D1Col12" nowrap><%=Trim(ObjRS("FullName"))%><BR></td>
											<td id="D1Col13" nowrap><%=Trim(ObjRS("TruckerCode"))%><BR></td>								
											<%
											If Trim(ObjRS("SenderCode1")) = USER AND Trim(ObjRS("TruckerCode"))<>COMPcd AND Trim(ObjRS("TruckerCode"))<>""  Then
												'指示先照会済みフラグ
												If ObjRS("TruckerFlag")=0 Then
													TruckerSubCode = "未"
												ElseIf ObjRS("TruckerFlag")=1 Then
													TruckerSubCode = "Yes"
												Else
													TruckerSubCode = "No"
												End If									
											Else
												TruckerSubCode = "　"
											End If
											
											%>
											<td id="D1Col14" nowrap><%=TruckerSubCode%><BR></td>	
											<% v_ItemName = "SenderCodeTbl1" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS("SenderCode1"))%>">
											<% v_ItemName = "BookNoTbl1" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS("BookNo"))%>">
											<% v_ItemName = "TruckerCodeTbl1" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS("TruckerCode"))%>">
										</tr>
										<% 
												ObjRS.MoveNext 												    
												End If
											Next							
											
											ObjRS.close
											DisConnDBH ObjConn, ObjRS
										%>													    						
									</tbody>
								</table>					
								<INPUT type=hidden name="DataCnt1" value="<%=x%>">					
							<% Else %>
								<table border="1" cellPadding="2" cellSpacing="0">						
								<TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
								</table>
							<% End If %>
							<!--Y.TAKAKUWA Add-S 2015-03-06-->
							<SCRIPT Language="JavaScript">
							cloneTable("TBEmpty", "TBEmpty1","1")
							</SCRIPT>
							<!--Y.TAKAKUWA Add-E 2015-03-06-->				
						</div>
						</td>
						</tr>
						</table>			
					</td>
				</tr>
				<tr>
					<td height="5"></td>
				</tr>
				<tr>		
					<td>	
						<table cellpadding="0" cellspacing="0">
							<td width="20"></td>
							<td width="200" nowrap><B>実搬出入事前情報（作業未完了）</B></td>
							<td>
							<select name="cmbINOut" onChange="fInOut();">
								<% If v_InOutFlag="1" Then %>
									<OPTION VALUE = ''>&nbsp;</OPTION>
									<OPTION VALUE = '1' SELECTED>実搬出</OPTION>
									<OPTION VALUE = '2'>実搬入</OPTION>
								<% ELSEIf v_InOutFlag="2" Then %>
									<OPTION VALUE = ''>&nbsp;</OPTION>
									<OPTION VALUE = '1'>実搬出</OPTION>
									<OPTION VALUE = '2' SELECTED>実搬入</OPTION>
								<% ELSE  %>
									<OPTION VALUE = '' SELECTED>&nbsp;</OPTION>
									<OPTION VALUE = '1'>実搬出</OPTION>
									<OPTION VALUE = '2'>実搬入</OPTION>
								<% END IF  %>
							</Select>
							<td width="50">&nbsp;</td>
							<td><input type="button" value="表示列設定" onClick="OpenCodeWin('表示列選択（実搬出入）','2')"></td>
							<td width="25">&nbsp;</td>
							<td><input type="button" value="並び替え"  onClick="OpenCodeWin2('6');"></td>				
							<td width="40">&nbsp;</td>
							<!--td width="100">&nbsp;</td-->
							<td>
								<%
									If Num2 > 0 Then
										abspage = ObjRS2.AbsolutePage
										pagecnt = ObjRS2.PageCount
										call LfPutPage(Num2,abspage,pagecnt,"pagenum2")
																	
										' Response.Write "<div align=""center"">" & vbcrlf
										' Response.Write "<a href="""
										' Response.Write Request.ServerVariables("SCRIPT_NAME")
										' Response.Write "?pagenum2=1""><b>最初のページ</b></a>"
										' Response.Write "	|	"
										'							
										' If abspage = 1 Then
										'	 Response.Write "<span>" & abspage & "</span>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 4 & """>&nbsp;<b>" & abspage + 4 & "</b></a>"
										' Elseif abspage = 2 then
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"							
										'	 Response.Write "<span>&nbsp;" & abspage & "</span>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
										' Elseif abspage = pagecnt then
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 4 & """>&nbsp;<b>" & abspage - 4 & "</b></a>"															
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"															
										'	 Response.Write "<span>&nbsp;" & abspage & "</span>"
										' Elseif abspage = CInt(pagecnt-1) then
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"															
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
										'	 Response.Write "<span>&nbsp;" & abspage & "</span>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"									
										' Else
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
										'	 Response.Write "<span>&nbsp;" & abspage & "</span>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
										'	 Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										'	 Response.Write "?pagenum2=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
										' End If
										'
										' Response.Write "	|	"
										' Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
										' Response.Write "?pagenum2=" & pagecnt & """><b>最後のページ</b></a>"
										' Response.Write "</div>" & vbcrlf
									End If
								%>
							</td>	
						</table>			
					</td>
				</tr>
				<tr>
					<td>
						<!--Y.TAKAKUWA Add-S 2015-03-05-->	
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<tr>
						<td>
							<DIV ID="BDIV2HEADER" style="overflow:hidden;">
							<table border="1" cellpadding="0" cellspacing="0" width="100%" Id="TBInOut1">										
							
							</table>
							</DIV>
						</td>
						</tr>
						<tr>
						<td>
						<!--Y.TAKAKUWA Add-E 2015-03-05-->
						<div id="BDIV2" onscroll="onScrollDiv(this,'BDIV2HEADER');">				
							<% If Num2>0 Then%>
								<% If blnSorted2 Then%>
								<!--<iframe frameborder="0" style="background-color:transparent;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400px;height:47px;"></iframe>-->				
								<% Else%>
								<!--<iframe frameborder="0" style="background-color:transparent;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400px;height:37px;"></iframe>-->									
								<% End If%>
								<table border="1" cellpadding="0" cellspacing="0" width="100%" Id="TBInOut">				
									<thead>
										<tr>
											<!--Y.TAKAKUWA Upd-S 2013-02-20-->
											<th id="H2Col01" class="hlist" align="left" nowrap><%=HeaderTbl2(0)%></th>	
											<!--2013-02-18 Y.TAKAKUWA ADD-S-->
											
											<th id="H2Col02" class="hlist" nowrap><%=HeaderTbl2(17)%></th>	
											<th id="H2Col03" class="hlist" nowrap><%=HeaderTbl2(18)%></th>
											
											<!--2013-02-18 Y.TAKAKUWA ADD-E-->						
											<th id="H2Col04" class="hlist" nowrap><%=HeaderTbl2(1)%></th>
											<th id="H2Col05" class="hlist" nowrap><%=HeaderTbl2(2)%></th>
											<th id="H2Col06" class="hlist" nowrap><%=HeaderTbl2(3)%></th>
											<th id="H2Col07" class="hlist" nowrap><%=HeaderTbl2(4)%></th>									
											<th id="H2Col08" class="hlist" nowrap><%=HeaderTbl2(5)%></th>
											<th id="H2Col09" class="hlist" nowrap><%=HeaderTbl2(6)%></th>
											<th id="H2Col10" class="hlist" nowrap><%=HeaderTbl2(7)%></th>
											<th id="H2Col11" class="hlist" nowrap><%=HeaderTbl2(8)%></th>
											<th id="H2Col12" class="hlist" nowrap><%=HeaderTbl2(9)%></th>
											<th id="H2Col13" class="hlist" nowrap><%=HeaderTbl2(10)%></th>
											<th id="H2Col14" class="hlist" nowrap><%=HeaderTbl2(11)%></th>
											<th id="H2Col15" class="hlist" nowrap><%=HeaderTbl2(12)%></th>
											<th id="H2Col16" class="hlist" nowrap><%=HeaderTbl2(13)%></th>
											<th id="H2Col17" class="hlist" nowrap><%=HeaderTbl2(14)%></th>
											<!--2010/05/12 Add-S C.Pestano-->
											<th id="H2Col18" class="hlist" nowrap><%=HeaderTbl2(15)%></th>
											<th id="H2Col19" class="hlist" nowrap><%=HeaderTbl2(16)%></th>						
											<!--2010/05/12 Add-E C.Pestano-->
											<!--Y.TAKAKUWA Upd-S 2013-02-20-->
										</tr>
									</thead>																	
									<tbody>
										<% 
											x = 1
											'While Not ObjRS2.EOF 							
											For i=1 To ObjRS2.PageSize
												If Not ObjRS2.EOF Then
												x = x + 1
										%>
										<% v_ItemName = "chkInOut" + cstr(i) %>
										<% if Trim(ObjRS2("Type")) = "FULLOUT" then%>
											<tr bgcolor="#FFCCFF">
											<td id="D2Col01" align="center" width="50" nowrap><input type="checkbox" name="<%= v_ItemName %>" disabled ><BR></td>
										<% else%>
											<tr bgcolor="#CCFFFF">
											<td id="D2Col01" align="center" width="50" nowrap><input type="checkbox" name="<%= v_ItemName %>"><BR></td>
										<% end if%>																	
											<!--2013-02-18 Y.TAKAKUWA ADD-S-->
											
											<%If Trim(ObjRS2("GroupFlag")) = "0" then %>
												<td id="D2Col02" nowrap style="color:Red;"><%=Trim(ObjRS2("LoDriverName"))%><BR></td>
												<td id="D2Col03" nowrap style="color:Red;">登録無<BR></td>
											<%else%>
												<td id="D2Col02" nowrap><%=Trim(ObjRS2("LoDriverName"))%><BR></td>
												<td id="D2Col03" nowrap><%=Trim(ObjRS2("LoHeadID"))%><BR></td>
											<%end if%>
											
											<!--2013-02-18 Y.TAKAKUWA ADD-E-->
											<td id="D2Col04" nowrap><%=Trim(ObjRS2("WorkDate"))%><BR></td>							
											<td id="D2Col05" nowrap><%=Trim(ObjRS2("Code1"))%><BR></td>
											<td id="D2Col06" nowrap><%=Trim(ObjRS2("Flag2"))%><BR></td>
											<% v_ItemName = "chkAns2No" + cstr(i) %>
											<td id="D2Col07" nowrap>								
											<select name= "<%= v_ItemName %>" class=chr>
											<option value="0">未</option>
											<option value="1">Yes</option>
											<option value="2">No</option>										
											</select>								
											</td>
											<%
												If Trim(ObjRS2("TruckerSubCode4")) = COMPcd Then
													TruckerSubCode = 4
												ElseIf Trim(ObjRS2("TruckerSubCode3")) = COMPcd Then
													TruckerSubCode = 3
												ElseIf Trim(ObjRS2("TruckerSubCode2")) = COMPcd Then
													TruckerSubCode = 2
												ElseIf Trim(ObjRS2("TruckerSubCode1")) = COMPcd Then
													TruckerSubCode = 1
												Else
													TruckerSubCode = 0
												end if										
											%>									
											<INPUT type=hidden name='Datatbl<%=i%>' value='<%=Trim(ObjRS2("DeliverTo"))%>,<%=Mid(ObjRS2("WorkDate"),3,8)%>,<%=Trim(ObjRS2("Code1"))%>,<%=Trim(ObjRS2("WkNo"))%>,<%=Trim(ObjRS2("FullOutType"))%>,<%=Trim(ObjRS2("BLContNo"))%>,<%=Trim(Mid(ObjRS2("WorkCompleteDate"),3,14))%>,<%=Trim(ObjRS2("ReturnDateStr"))%>,<%=ObjRS2("ReturnValue")%>,<%=Trim(ObjRS2("Code2"))%>,<%=ObjRS2("Flag1")%>,<%=Trim(ObjRS2("BLNo"))%>,<%="-"%>,<%=TruckerSubCode%>,<%=Trim(ObjRS2("Flag2"))%>,<%=Trim(ObjRS2("ShipLine"))%>,<%=Trim(ObjRS2("ShipName"))%>,<%=Trim(ObjRS2("ContSize"))%>,<%=Trim(ObjRS2("CY"))%>,<%=Mid(ObjRS2("FreeTime"),3,8)%>,<%=TruckerSubCode%>,<%=Trim(ObjRS2("ReturnDateVal"))%>,<%=Trim(ObjRS2("Comment1"))%>,<%=Trim(ObjRS2("Comment2"))%>,<%=Trim(ObjRS2("DeliverTo1"))%>,<%=Left(Trim(ObjRS2("Name1")),8)%>'>
											<% if Trim(ObjRS2("Type")) = "FULLOUT" then%>						
											<td id="D2Col08" nowrap><A HREF="JavaScript:GoRenew('<%=i%>');"><%=Trim(ObjRS2("WkNo"))%></A><BR></td>																
											<% else %>
											<td id="D2Col08" nowrap><A HREF="JavaScript:GoRenew2('<%=ObjRS2("WkNo")%>','<%=ObjRS2("BookNo")%>','<%=ObjRS2("ContNo")%>');"><%=Trim(ObjRS2("WkNo"))%></A><BR></td>																
											<% end if %>
											<% if Trim(ObjRS2("Type")) = "FULLOUT" then%>						
											<td id="D2Col09" nowrap><A HREF="JavaScript:GoConinf2('<%=Trim(ObjRS2("WkNo"))%>','<%=Trim(ObjRS2("FullOutType"))%>','<%=Trim(ObjRS2("BLContNo"))%>')"><%=Trim(ObjRS2("BLContNo"))%></A><BR></td>																
											<% else %>
											<td id="D2Col09" nowrap><A HREF="JavaScript:GoConinfContNo('<%=Trim(ObjRS2("BLContNo"))%>');"><%=Trim(ObjRS2("BLContNo"))%></A><BR></td>																
											<% end if %>
											<td id="D2Col10" nowrap><%=Trim(ObjRS2("ShipLine"))%><BR></td>
											<td id="D2Col11" nowrap><%=Trim(ObjRS2("ShipName"))%><BR></td>
											<td id="D2Col12" nowrap><%=Trim(ObjRS2("ContSize"))%><BR></td>
											<td id="D2Col13" nowrap><%=Trim(ObjRS2("ReceiveFrom"))%><BR></td>
											<td id="D2Col14" nowrap><%=Trim(ObjRS2("CY"))%><BR></td>
											<% if Trim(ObjRS2("DelPermitDate")) = "-" then%>
												<td id="D2Col15" nowrap align="center"><%=Trim(ObjRS2("DelPermitDate"))%><BR></td>
											<% else%>
												<td id="D2Col15" nowrap><%=Trim(ObjRS2("DelPermitDate"))%><BR></td>
											<% end if%>								
											<td id="D2Col16" nowrap><%=Trim(ObjRS2("FreeTime"))%><BR></td>								
											<td id="D2Col17" nowrap><%=Trim(ObjRS2("CYCut"))%><BR></td>
											<!--2010/05/12 Add-S C.Pestano-->
											<td id="D2Col18" nowrap><%=Trim(ObjRS2("Code2"))%><BR></td>
											<td id="D2Col19" nowrap><%=Trim(ObjRS2("Flag1"))%><BR></td>
											<!--2010/05/12 Add-E C.Pestano-->
											<% v_ItemName = "WkNo" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("WkNo"))%>">
											<% v_ItemName = "ContNo" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("BLContNo"))%>">
											<% v_ItemName = "BookNo" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("BookNo"))%>">	
											<% v_ItemName = "Flag2" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("Flag2"))%>">
											<% v_ItemName = "TruckerSubCodeTbl2" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=TruckerSubCode%>">																																
											<% v_ItemName = "TypeCode" + cstr(i) %>
											<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS2("Type"))%>">
										</tr>
									<% 
												ObjRS2.MoveNext 		
												End If
											Next	
										'Wend
									
										ObjRS2.close    
										DisConnDBH ObjConn2, ObjRS2
							
									%>						    									
									</tbody>								
								</table>
								<INPUT type=hidden name="DataCnt2" value="<%=x%>">
							<% Else %>
								<table border="1" cellPadding="2" cellSpacing="0">						
									<TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
								</table>
							<% End If%>
							<!--Y.TAKAKUWA Add-S 2015-03-06-->
							<SCRIPT Language="JavaScript">
							cloneTable("TBInOut", "TBInOut1","2")
							</SCRIPT>
							<!--Y.TAKAKUWA Add-E 2015-03-06-->		
						</div>
						</td>
						</tr>
						</table>
					</td>	
				</tr>
				<tr>
					<td height="5"></td>
				</tr>
				<tr>
					<td>
						<table border="0" cellpadding="2" cellspacing="0">
						<tr>
						<td><input type="button" name="btnPrint"  value="搬入票出力" onClick="fPrint();"></td>
						<td><input type="button" name="btnUpdate" value="回答する" onClick="UpdAns();"></td>
						</tr>
						</table>
					</td>
				</tr>	
			</table>
		</div>
	</form>
	<!-- 2010/04/27-2 Upd-S C.Pestano-->
<!--/div-->
</BODY>
</HTML>
