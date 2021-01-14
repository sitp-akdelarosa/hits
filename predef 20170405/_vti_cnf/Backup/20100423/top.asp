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
<%'**********************************************
  '共通の前提処理
  '共通関数  (Commonfunc.inc)
%>
<!--#include File="Common.inc"-->
<!--#include File="../download/download.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->
<%
	'**********************************************
	'ユーザデータ所得
	dim USER, COMPcd  			
	dim v_GamenMode
	
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
	dim i
	dim v_ItemName
	dim v_loop
	dim v_work
	dim file1,gerrmsg 
	dim Arr_SCACCode()
	dim v_SCACCode
	dim strSort
	dim abspage, pagecnt	
	
	'INIファイルより設定値を取得
	dim param(2),calcDate
	getIni param
  	calcDate = DateAdd("d", "-"&param(2), Date)
  
	ReDim FieldName1(13)	
	FieldName1=Array("InPutDate","SenderCode","Flag2","","BookNo","NumCount","ContSize1","ContType1","ContHeight1","ContMaterial1","ShipLine","FullName","TruckerCode","TruckerFlag")  		
			
	ReDim FieldName2(12)	
	FieldName2=Array("WorkDate","Code1","Flag2","WkNo","BLContNo","ShipLine","ShipName","ContSize","ReceiveFrom","CY","DelPermitDate","FreeTime","CYCut")    
	
	Const CONST_ASC = "<BR><IMG border=0 src=Image/ascending.gif>"
	Const CONST_DESC = "<BR><IMG border=0 src=Image/descending.gif>"
	
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
		if Mid(strSort,1,1) = "X" then
			Session("Key1") = ""
		else
			Session("Key1") = Mid(strSort,1,1)
		end if
		
		if Mid(strSort,2,1) = "0" then
			Session("KeySort1") = "ASC"
		else
			Session("KeySort1") = "DESC"
		end if	
		
		if Mid(strSort,3,1) = "X" then
			Session("Key2") = ""
		else
			Session("Key2") = Mid(strSort,3,1)
		end if
		
		if Mid(strSort,4,1) = "0" then
			Session("KeySort2") = "ASC"
		else
			Session("KeySort2") = "DESC"
		end if			
		
		if Mid(strSort,5,1) = "X" then
			Session("Key3") = ""
		else
			Session("Key3") = Mid(strSort,5,1)
		end if
				
		if Mid(strSort,6,1) = "0" then
			Session("KeySort3") = "ASC"
		else
			Session("KeySort3") = "DESC"
		end if	
	end if
	
	strOrder1 = getSort1(Session("Key1"),Session("KeySort1"),"")
	strOrder1 = getSort1(Session("Key2"),Session("KeySort2"),strOrder1)
	strOrder1 = getSort1(Session("Key3"),Session("KeySort3"),strOrder1)		
	
	strSort = request.cookies("SortTbl2")	
	
	if strSort <> "" then
		if Mid(strSort,1,1) = "X" then
			Session("TB2Key1") = ""
		else
			Session("TB2Key1") = Mid(strSort,1,1)
		end if
		
		if Mid(strSort,2,1) = "0" then
			Session("TB2KeySort1") = "ASC"
		else
			Session("TB2KeySort1") = "DESC"
		end if	
		
		if Mid(strSort,3,1) = "X" then
			Session("TB2Key2") = ""
		else
			Session("TB2Key2") = Mid(strSort,3,1)
		end if
		
		if Mid(strSort,4,1) = "0" then
			Session("TB2KeySort2") = "ASC"
		else
			Session("TB2KeySort2") = "DESC"
		end if			
		
		if Mid(strSort,5,1) = "X" then
			Session("TB2Key3") = ""
		else
			Session("TB2Key3") = Mid(strSort,5,1)
		end if
				
		if Mid(strSort,6,1) = "0" then
			Session("TB2KeySort3") = "ASC"
		else
			Session("TB2KeySort3") = "DESC"
		end if	
	end if
	
	strOrder2 = getSort2(Session("TB2Key1"),Session("TB2KeySort1"),"")
	strOrder2 = getSort2(Session("TB2Key2"),Session("TB2KeySort2"),strOrder2)
	strOrder2 = getSort2(Session("TB2Key3"),Session("TB2KeySort3"),strOrder2)
	
	if v_SCACCode <> "" then
		strWhere1 = strWhere1 & " AND T.ShipLine = '" & Trim(v_SCACCode) & "'"
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

Function LfGetRequestItem()
	v_GamenMode = Request.form("Gamen_Mode") 
	v_SCACCode = Request.form("cmbSCACCode") 
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
	HeaderTbl1(6) = "サイズ"
	HeaderTbl1(7) = "タイプ"
	HeaderTbl1(8) = "高さ"
	HeaderTbl1(9) = "材質"
	HeaderTbl1(10) = "船社"
	HeaderTbl1(11) = "船名"
	HeaderTbl1(12) = "指示先"
	HeaderTbl1(13) = "指示先<BR>回答"
		 
	for ctr = 1 to 3	
		Session(CSTR("Key" & ctr))	
		if Session(CSTR("Key" & ctr)) <> "" then
			Select Case Session(CSTR("Key" & ctr))
				Case "0" '入力日
					HeaderTbl1(0) = HeaderTbl1(0) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "1" '指示元
					HeaderTbl1(1) = HeaderTbl1(1) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "2" '指示元へ回答
					HeaderTbl1(2) = HeaderTbl1(2) & getImage(Session(CSTR("KeySort" & ctr)))
					blnSorted = True				
				Case "4" 'ブッキング番号
					HeaderTbl1(4) = HeaderTbl1(4) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "5" 'ピック済
					HeaderTbl1(5) = HeaderTbl1(5) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "6" 'サイズ
					HeaderTbl1(6) = HeaderTbl1(6) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "7" 'タイプ
					HeaderTbl1(7) = HeaderTbl1(7) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "8" '高さ
					HeaderTbl1(8) = HeaderTbl1(8) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "9" '材質
					HeaderTbl1(9) = HeaderTbl1(9) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "10" '船社
					HeaderTbl1(10) = HeaderTbl1(10) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "11" '船名
					HeaderTbl1(11) = HeaderTbl1(11) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "12" '指示先
					HeaderTbl1(12) = HeaderTbl1(12) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "13" '指示先へ回答
					HeaderTbl1(13) = HeaderTbl1(13) & getImage(Session(CSTR("KeySort" & ctr)))	
					blnSorted = True							
			  End Select
		end if	  
	next
		
	'" LEFT JOIN Booking as BOK on VECO.VslCode=BOK.VslCode and VECO.VoyCtrl=BOK.VoyCtrl and VECO.BookNo=BOK.BookNo"&_
	StrSQL = "SELECT * FROM(SELECT DISTINCT SPB.BookNo, mV.FullName, SPB.SenderCode," &_
			 " ISNULL(CONVERT(varchar(10),SPB.InputDate,111),'') as InputDate, SPB.TruckerCode, SPB.TruckerFlag,"&_
			 " (CASE SPB.TruckerFlag WHEN 0 THEN '未' WHEN 1 THEN 'Yes' WHEN 2 THEN 'No' ELSE ' ' END) as Flag2,"&_ 
			 " SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1, mS.OpeCode,"&_
			 " (CASE WHEN mU.UserType = '5' THEN mU.HeadCompanyCode ELSE SPB.SenderCode END) as Code1, "&_
			 " (CASE WHEN mU.UserType = '5' THEN mU.TTName ELSE mU.TTName END) as TruckerName, "&_
			 " IsNull(CASE (VPC.Picks) WHEN '1' THEN VPC.PickPlace ELSE '複数' END ,'') PickPlace,"&_
			 " IsNull(VEC.numC,'') as NumCount,"&_
			 " SPB.Comment1,SPB.Comment2, mU.HeadCompanyCode, mU.UserType ,SPB.ShipLine,VSLS.CYCut "&_
			 " FROM BookingAssign AS SPB "&_
			 " LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_
			 " LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO = VBA.BookNo"&_
			 " LEFT JOIN ViewExportCnt As VEC ON SPB.BookNo = VEC.BookNo"&_
			 " LEFT JOIN ViewExportCont As VECO ON SPB.BookNo = VECO.BookNo"&_	
			 		 
			 
			 
			 " LEFT JOIN (select a.bookno bookno ,b.vslcode vslcode , b.voyctrl voyctrl , "&_
			 " ISNULL(a.shipline,b.shipline) shipline "&_
			 " FROM bookingassign A left join booking b on a.bookno=b.bookno "&_
			 " WHERE (A.SenderCode='"& USER &"' OR A.TruckerCode='"& COMPcd &"') AND A.Process='R' "&_
			 " AND DateDiff(day,A.InputDate,'"&calcDate&"')<=0" &_
			 ") as BOK on VECO.VslCode=BOK.VslCode and VECO.VoyCtrl=BOK.VoyCtrl and VECO.BookNo=BOK.BookNo "&_			   
			 
			 
			 
			 " LEFT JOIN VslSchedule AS VSLS ON BOK.VslCode = VSLS.VslCode AND BOK.VoyCtrl = VSLS.VoyCtrl"&_ 
			 " LEFT JOIN mVessel AS mV ON BOK.VslCode = mV.VslCode"&_
			 " LEFT JOIN mShipLine AS mS ON SPB.ShipLine = mS.ShipLine"&_ 
			 " LEFT JOIN ViewPickupCnt AS VPC ON BOK.VslCode = VPC.VslCode AND BOK.VoyCtrl = VPC.VoyCtrl AND BOK.BookNo = VPC.BookNo"&_ 
			 " WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R'"&_
			 " AND DateDiff(day,SPB.InputDate,'"&calcDate&"')<=0" &_
			 " ) AS T"&_
			 " WHERE T.Flag2='未' OR T.Flag2=' ' "&_			 
			 strWhere & strOrder1		
	
	ObjRS.PageSize = 100	
	ObjRS.CacheSize = 100
	ObjRS.CursorLocation = 3				
	ObjRS.Open StrSQL, ObjConn	
	Num = ObjRS.recordcount		
	
	if Num > 100 then
		If Len(Request("pagenum")) = 0 Then
			ObjRS.AbsolutePage = 1
		Else
			If CInt(Request("pagenum")) <= ObjRS.PageCount Then
				ObjRS.AbsolutePage = Request("pagenum")
			Else
				ObjRS.AbsolutePage = 1
			End If
		End If
	End If
	 
	if err <> 0 then			
		DisConnDBH ObjConn, ObjRS	'DB切断
		jampErrerP "2","b301","01","空搬出事前情報一覧","102","SQL:<BR>" & strSQL
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
	
	On Error Resume Next

	ConnDBH ObjConn2, ObjRS2	
	
	WriteLogH "b401", "実搬入事前情報一覧", "01", ""
	
	ReDim HeaderTbl2(14)	
	
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
		
	for ctr = 1 to 3	
		Session(CSTR("TB2Key" & ctr))	
		if Session(CSTR("TB2Key" & ctr)) <> "" then
			Select Case Session(CSTR("TB2Key" & ctr))
				Case "0" '搬出入予定日
					HeaderTbl2(1) = HeaderTbl2(1) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "1" '指示元
					HeaderTbl2(2) = HeaderTbl2(2) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "2" '指示元へ回答
					HeaderTbl2(3) = HeaderTbl2(3) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "3" '作業番号
					HeaderTbl2(5) = HeaderTbl2(5) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "4" 'コンテナ番号/BL番号
					HeaderTbl2(6) = HeaderTbl2(6) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "5" '船社
					HeaderTbl2(7) = HeaderTbl2(7) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "6" '船名
					HeaderTbl2(8) = HeaderTbl2(8) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "7" 'SZ
					HeaderTbl2(9) = HeaderTbl2(9) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "8" '搬入元
					HeaderTbl2(10) = HeaderTbl2(10) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "9" 'CY
					HeaderTbl2(11) = HeaderTbl2(11) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "10" '搬出可否
					HeaderTbl2(12) = HeaderTbl2(12) & getImage(Session(CSTR("TB2KeySort" & ctr)))
				Case "11" 'フリータイム
					HeaderTbl2(13) = HeaderTbl2(13) & getImage(Session(CSTR("TB2KeySort" & ctr)))
					blnSorted2 = True
				Case "12" 'CYカット日
					HeaderTbl2(14) = HeaderTbl2(14) & getImage(Session(CSTR("TB2KeySort" & ctr)))			
			End Select
		end if	  
	next					
		
'	StrSQL ="SELECT  * FROM (SELECT T.* FROM (SELECT 'FULLOUT' As Type,ITC.DeliverTo,ITC.BLNo, " & _
'			"ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate, "&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
'			"ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
'			"END) as Code1, "&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
'			"ELSE ITC.TruckerSubName1 "&_
'			"END) as Name1, "&_
'			"(CASE "&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
'			"ELSE Null END) "&_
'			"WHEN 0 THEN '未' "&_
'			"WHEN 1 THEN 'Yes' "&_
'			"WHEN 2 THEN 'No' "&_
'			"ELSE ' ' END) as Flag2, "&_
'			"ITC.WkNo, "&_
'			"ITC.FullOutType, "&_
'			"ITC.BLNo as BLContNo, A.ShipLine, A.FullName as ShipName, ''  as ContSize, "&_
'			"SUBSTRING(A.RecTerminal,1,2) as CY, "&_
'			"ISNULL(CONVERT(varchar(10),A.FreeTime,111),'') as FreeTime, "&_
'			"ITC.DeliverTo1,ITC.WorkCompleteDate, "&_
'			"ITC.ReturnDateStr, "&_
'			"(CASE WHEN ITC.FullOutType = '1' THEN (CASE WHEN INC.ReturnTime IS NULL THEN '未' ELSE '済' END) "&_
'			"ELSE ' ' END) as ReturnValue, "&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
'			"ELSE ITC.TruckerSubCode1 "&_
'			"END) as Code2, "&_
'			"(CASE WHEN "&_
'			"(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
'			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
'			"ELSE ITC.TruckerSubCode1 "&_
'			"END) IS NULL THEN ' ' ELSE "&_
'			"(CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
'			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
'			"ELSE ITR.TruckerFlag1 "&_
'			"END) "&_
'			"WHEN '0' THEN '未' "&_
'			"WHEN '1' THEN 'Yes' "&_
'			"ELSE 'No' END) "&_
'			"END) as Flag1, "&_
'			"ITC.Comment1, ITC.Comment2, "&_
'			"ITC.ReturnDateVal, ITC.UpdtUserCode, "&_
'			"ITR.TruckerFlag1,ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
'			"mU.HeadCompanyCode, mU.UserType, "&_
'			"A.CYDelTime, INC.ReturnTime,ITC.InputDate, "&_
'			"'' as InputDate1,'' as WorkDate1,'' as BookNo,'' as ContNo,'' as VslName,'' as ContHeight,'' as TareWeight, "&_
'			"'' as ReceiveFrom,ISNULL(CONVERT(varchar(10),INC.DelPermitDate,111),'-') as DelPermitDate,'' as CYCut,'' as CYCut1,'' as WorkComplete,'' as WorkComplete1, "&_
'			"ITC.TruckerSubCode1,ITC.TruckerSubCode2,ITC.TruckerSubCode3,ITC.TruckerSubCode4,ITC.WkContrlNo, "&_
'			"'' as Nine , '' as Comment3,'' as RegisterCode "&_			
'			"FROM hITCommonInfo ITC "&_
'			"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
'			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
'			"LEFT JOIN ImportCont AS INC ON (ITC.ContNo=INC.ContNo) "&_
'			"LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode) "&_
'			"LEFT JOIN (SELECT Distinct BL.BLNo,INC.FreeTime, "&_
'			"MIN(INC.CYDelTime) AS CYDelTime,mV.ShipLine, mV.FullName, BL.RecTerminal "&_
'			"FROM ImportCont AS INC "&_
'			"LEFT JOIN mVessel AS mV ON INC.VslCode = mV.VslCode "&_
'			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
'			"GROUP BY BL.BLNo,INC.FreeTime,mV.ShipLine, mV.FullName, BL.RecTerminal) A ON  A.BLNo=ITC.BLNo "&_
'			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType='4' AND (ITC.RegisterCode='"& USER &"' "&_
'			"OR ITC.TruckerSubCode1='"& COMPcd &"' "&_
'			"OR ITC.TruckerSubCode2='"& COMPcd &"' "&_
'			"OR ITC.TruckerSubCode3='"& COMPcd &"' "&_
'			"OR ITC.TruckerSubCode4='"& COMPcd &"') "&_
'			"UNION ALL "&_
'			"SELECT 'FULLOUT' As Type,ITC.DeliverTo,ITC.BLNo, "&_
'			"ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate,"&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
'			"ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) "&_
'			"END) as Code1, "&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
'			"ELSE ITC.TruckerSubName1 "&_
'			"END) as Name1, "&_
'			"(CASE "&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
'			"ELSE Null END) "&_
'			"WHEN 0 THEN '未' "&_
'			"WHEN 1 THEN 'Yes' "&_
'			"WHEN 2 THEN 'No' "&_
'			"ELSE ' ' END) as Flag2, "&_
'			"ITC.WkNo, "&_
'			"ITC.FullOutType, "&_
'			"ITC.ContNo as BLContNo, "&_
'			"mV.ShipLine, "&_
'			"mV.FullName as ShipName, "&_
'			"CON.ContSize, "&_
'			"SUBSTRING(BL.RecTerminal,1,2) as CY, "&_
'			"ISNULL(CONVERT(varchar(10),INC.FreeTime,111),'') as FreeTime, "&_
'			"ITC.DeliverTo1, "&_
'			"ITC.WorkCompleteDate, "&_
'			"ITC.ReturnDateStr, "&_
'			"(CASE WHEN ITC.FullOutType = '1' THEN (CASE WHEN INC.ReturnTime IS NULL THEN '未' "&_
'			" ELSE '済' END) "&_
'			"ELSE ' ' END) as ReturnValue, "&_
'			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
'			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
'			"ELSE ITC.TruckerSubCode1 "&_
'			"END) as Code2, "&_
'			"(CASE WHEN "&_
'			"(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
'			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
'			"ELSE ITC.TruckerSubCode1 "&_
'			"END) IS NULL THEN ' ' "&_
'			"ELSE "&_
'			"(CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
'			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
'			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
'			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
'			"ELSE ITR.TruckerFlag1 "&_
'			"END) "&_
'			"WHEN '0' THEN '未' WHEN '1' THEN 'Yes' ELSE 'No' END) END) as Flag1, "&_
'			"ITC.Comment1, ITC.Comment2, "&_
'			"ITC.ReturnDateVal, ITC.UpdtUserCode, "&_
'			"ITR.TruckerFlag1,ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, "&_
'			"mU.HeadCompanyCode, mU.UserType, "&_
'			"INC.CYDelTime, INC.ReturnTime,ITC.InputDate, "&_
'			"'' as InputDate1,'' as WorkDate1,'' as BookNo,'' as ContNo,'' as VslName,'' as ContHeight,'' as TareWeight, "&_
'			"'' as ReceiveFrom,ISNULL(CONVERT(varchar(10),INC.DelPermitDate,111),'-') as DelPermitDate,'' as CYCut,'' as CYCut1,'' as WorkComplete,'' as WorkComplete1, "&_
'			"ITC.TruckerSubCode1,ITC.TruckerSubCode2,ITC.TruckerSubCode3,ITC.TruckerSubCode4,ITC.WkContrlNo, "&_
'			"'' as Nine , '' as Comment3,'' as RegisterCode "&_	
'			"FROM hITCommonInfo ITC "&_
'			"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
'			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
'			"LEFT JOIN ImportCont AS INC ON ITC.ContNo=INC.ContNo "&_
'			"LEFT JOIN mVessel AS mV On INC.VslCode=mV.VslCode "&_
'			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
'			"LEFT JOIN Container AS CON ON INC.VslCode = CON.VslCode AND INC.VoyCtrl=CON.VoyCtrl AND INC.ContNo=CON.ContNo "&_
'			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType<>'4' AND (ITC.RegisterCode='"& USER &"' "&_
'			"OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"')) AS T " &_
'            "UNION ALL " &_			
'			"SELECT T.* FROM (SELECT 'FULLIN' As Type,'' as DeliverTo, '' as BLNo, ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate," &_
'            "        (CASE " &_
'    	    "            WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) " &_
'            "            WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 " &_
'            "            WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
'            "            WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
'            "            ELSE (CASE WHEN mU.UserType = 5 THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END) " &_
'		    "        END) as Code1, " &_
'		    "		 (CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
'			"			WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
'			"			WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
'			"			WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
'			"			ELSE ITC.TruckerSubName1 "&_
'			"		 END) as Name1, "&_
'		    "		(CASE "&_
'			"		(CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_
'			"		WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_
'			"		WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
'			"		WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
'			"		ELSE Null END) "&_
'			"		WHEN 0 THEN '未' "&_
'			"		WHEN 1 THEN 'Yes' "&_
'			"		WHEN 2 THEN 'No' "&_
'			"		ELSE ' ' END) as Flag2, "&_			
'		    "       ITC.WkNo, '' as FullOutType,ITC.ContNo as BLContNo, CYV.ShipLine,'' as ShipName, " &_
'			"		CYV.ContSize,'' as CY, '' as FreeTime, '' as DeliverTo1, " &_			
'		    "       '' as WorkCompleteDate,'' as ReturnDateStr,'' as ReturnValue, " &_
'		    "        (CASE " &_
'			"            WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
'			"            WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 " &_
'			"            WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
'		    "            WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
'		    "            ELSE ITC.TruckerSubCode1 " &_
'		    "        END) as Code2, " &_
'		    "        (CASE WHEN " &_
'		    "              (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 " &_
'		    "                    WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 " &_
'		    "                    WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 " &_
'		    "                    WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
'		    "                    ELSE ITC.TruckerSubCode1 " &_
'		    "               END) IS NULL THEN ' '" &_
'		    "              ELSE " &_
'		    "                    (CASE (CASE WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 " &_
'		    "                                WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 " &_
'		    "                                WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 " &_
'		    "                                WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL " &_
'		    "                                ELSE ITR.TruckerFlag1 " &_
'		    "                     END) " &_
'		    "             WHEN '0' THEN '未' " &_
'		    "             WHEN '1' THEN 'Yes' " &_
'		    "             ELSE 'No' END) " &_
'		    "          END) as Flag1, " &_
'			"        SUBSTRING(ITC.Comment1,1,10) as Comment1, SUBSTRING(ITC.Comment2,1,10) as Comment2, " &_
'			" 		'' as ReturnDateVal, ITC.UpdtUserCode,ITR.TruckerFlag1, ITR.TruckerFlag2, ITR.TruckerFlag3, " &_
'			"		ITR.TruckerFlag4, mU.HeadCompanyCode, mU.UserType,'' as CYDelTime,'' as ReturnTime," &_
'			"		CONVERT(varchar,ITC.InputDate,11) as InputDate,ITC.InputDate as InputDate1, " &_ 
'			"		ITC.WorkDate as WorkDate1,CYV.BookNo,ITC.ContNo,CYV.VslName as VslName, " &_
'			" 		CYV.ContHeight,CASE ISNULL(CYV.TareWeight,0) WHEN 0 THEN '-' ELSE CYV.TareWeight END TareWeight, " &_
'			"		SUBSTRING(CYV.ReceiveFrom,1,20) as ReceiveFrom, '-' as DelPermitDate," &_
'			"		ISNULL(CONVERT(varchar(10),VSLS.CYCut,111),'') as CYCut, VSLS.CYCut as CYCut1, " &_			
'			"		CONVERT(varchar,ITC.WorkCompleteDate,11) + ' ' + Substring(CONVERT(varchar,ITC.WorkCompleteDate,8),1,5) as WorkComplete, " &_
'			"		ITC.WorkCompleteDate as WorkComplete1, " &_
'			"		ITC.TruckerSubCode1, ITC.TruckerSubCode2, ITC.TruckerSubCode3, ITC.TruckerSubCode4,  ITC.WkContrlNo, " & _
'			"          (CASE "&_
'			"                WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN '1' "&_
'			"                WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN '2' "&_
'			"                WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN '3' "&_
'			"                WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '4' "&_
'			"                ELSE '0' END) as Nine, "&_		
'			"        SUBSTRING(ITC.Comment3,1,10) as Comment3, " &_
'		    "        ITC.RegisterCode " &_  
'		    "        FROM hITCommonInfo AS ITC " &_
'		    "        INNER JOIN CYVanInfo AS CYV ON ITC.ContNo=CYV.ContNo AND ITC.WkNo = CYV.WkNo " &_
'		    "        INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo " &_
'		    "        INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode " &_
'		    "        LEFT JOIN ExportCont AS EPC ON CYV.ContNo = EPC.ContNo AND CYV.BookNo = EPC.BookNo and CYV.VslCode = EPC.VslCode " &_
'
'		    "        LEFT JOIN VslSchedule AS VSLS ON EPC.VslCode = VSLS.VslCode AND EPC.VoyCtrl = VSLS.VoyCtrl " &_
'		    "        LEFT JOIN Booking AS BOK ON EPC.VslCode = BOK.VslCode AND EPC.VoyCtrl = BOK.VoyCtrl AND EPC.BookNo = BOK.BookNo " &_
'		    "        WHERE ITC.Process='R' AND ITC.WkType='3' AND (ITC.RegisterCode='"& USER &"' " &_
'		    "        OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' " &_
'		    "        OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"')" &_
'			"        ) AS T ) AS A WHERE A.Flag2 = '未' OR A.Flag2 = ' ' " &_ 
'			strOrder2
'			
'	ObjRS2.PageSize = 100
'	ObjRS2.CacheSize = 100
'	ObjRS2.CursorLocation = 3
'	ObjRS2.Open StrSQL, ObjConn2
'	Num2 = ObjRS2.recordcount	
'
'	if Num2 > 100 then
'		If Len(Request("pagenum2")) = 0 Then
'			ObjRS2.AbsolutePage = 1
'		Else
'			If CInt(Request("pagenum2")) <= ObjRS2.PageCount Then
'				ObjRS2.AbsolutePage = Request("pagenum2")
'			Else
'				ObjRS2.AbsolutePage = 1
'			End If
'		End If		 
'	end if
'	
'	if err <> 0 then
'		DisConnDBH ObjConn2, ObjRS2	'DB切断
'		jampErrerP "2","b301","01","実搬出入事前情報","102","SQL:<BR>" & StrSQL & err.description
'		Exit Function
'	end if			

	'エラートラップ解除
    on error goto 0	
End Function

Function LfUpdTruckerAns()
	dim ObjConnUpd, ObjRSUpd, StrSQL
	
	'エラートラップ開始
	On Error Resume Next
	
	'DB接続	
	ConnDBH ObjConnUpd, ObjRSUpd

	For i=1 To ObjRS.PageSize
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
	
	For i=1 To ObjRS2.PageSize	
		if CInt(Trim(Request.form("TruckerSubCodeTbl2" & i))) > 0 AND Trim(Request.form("Flag2" & i)) = "未" AND Trim(Request.form("chkAns2No" & i)) <> "0" then  						
			StrSQL = "UPDATE hITReference SET TruckerFlag" & Trim(Request.form("TruckerSubCodeTbl2" & i)) & "='" & Trim(Request.form("chkAns2No" & i)) & "'," & _
				     " UpdtTime='"& Now() &"', UpdtPgCd='TOP',UpdtTmnl='" & USER & "'" & _
					 " WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo " & _
					 " WHERE WkNo='"& Trim(Request.form("WkNo" & i)) &"' AND WkType='1' AND Process='R')"		
						
			ObjConnUpd.Execute(StrSQL)
			
			if err <> 0 then
				Set ObjRSUpd = Nothing				
				jampErrerPDB ObjConnUpd,"2","b107","01","実搬出:紹介済処理","104","SQL:<BR>"&strSQL
			end if		
		end if		
	Next	
	'DB接続解除
	DisConnDBH ObjConnUpd, ObjRSUpd

	'エラートラップ解除
	on error goto 0		  
End Function

Function getSort1(Key,SortKey,str)
	getSort1 = str
	if Key <> "" then	
		if str = "" then
			getSort1 = " ORDER BY (Case When LTRIM(ISNULL(" & FieldName1(Key) & ",''))='' Then 1 Else 0 End), " & FieldName1(Key) & " " & SortKey
		else
			getSort1 = str & " , (Case When LTRIM(ISNULL(" & FieldName1(Key) & ",''))='' Then 1 Else 0 End), " & FieldName1(Key) & " " & SortKey
		end if	
	end if	
end function

Function getSort2(Key,SortKey,str)
	getSort2 = str
	if Key <> "" then	
		if str = "" then
			getSort2 = " ORDER BY (Case When LTRIM(ISNULL(" & FieldName2(Key) & ",''))='' Then 1 Else 0 End), " & FieldName2(Key) & " " & SortKey
		else
			getSort2 = str & " , (Case When LTRIM(ISNULL(" & FieldName2(Key) & ",''))='' Then 1 Else 0 End), " & FieldName2(Key) & " " & SortKey
		end if	
	end if	
end function

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
DIV.BDIV2 {
    position: relative;
    border-width: 0px 0px 1px 0px;
}
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
</style>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT Language="JavaScript">
function finit(){
	var str;
	
	//データ引継ぎ設定  
    document.frm.Gamen_Mode.value="<%=v_GamenMode%>";    //処理モード
	
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

//データが無い場合の表示制御
function view(){
	
	var obj1=document.getElementById("BDIV1");

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
		obj1.style.height = 150;
		obj1.style.overflowY = "auto";
	}		
	
	var obj2=document.getElementById("BDIV2");
	
	if((document.body.offsetWidth-200) < 50){
		obj2.style.width=50;
		obj2.style.overflowX="auto";	 
	}else if((document.body.offsetWidth-200)  < 813){
		//obj2.style.width=document.body.offsetWidth-200;
		obj2.style.width=document.body.offsetWidth-30;
		obj2.style.overflowX="auto";
	}else{
		//obj2.style.width=document.body.offsetWidth-200;
		obj2.style.width=document.body.offsetWidth-30;
		obj2.style.overflowX="auto";
	}	
	
	if((document.body.offsetHeight-280) < 50){
		obj2.style.height = 50;
		obj2.styleoverflowY = "auto";	 
	}else if((document.body.offsetHeight-280)  < 335){
		obj2.style.height = document.body.offsetHeight-280;
		obj2.style.overflowY = "auto";
	}else{
		obj2.style.height = document.body.offsetHeight-280;
		obj2.style.overflowY = "auto";
	}
	
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

    var w=900;
    var h=550;
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
	newWin = window.open("", "ReEntry", "left=10,top=10,status=yes,resizable=yes,scrollbars=yes");
	Fname.target="ReEntry";
	Fname.elements[No].disabled=false;
	Fname.submit();
	Fname.elements[No].disabled=true;
	Fname.target="_self";
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

</SCRIPT>
</HEAD>
<BODY onLoad="setTimeout('showContent()', 500);finit();view();" onResize="view();">
<div class="center" id="loading2">しばらくお待ちください。&nbsp;<IMG border=0 src=Image/loaded.gif></div>
<div id="content" style="display:none;">
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
				</tr>									
				<tr>
					<td width="20"></td>
					<td nowrap><B>空搬出事前情報</B></td>
					<td>
						<select name="cmbSCACCode" onChange="document.frm.submit();">
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
					<td><input name="btn2" type="button" value="ソート設定" onClick="OpenCodeWin2('5');"></td>
					<td><input name="btn3" type="button" value="更新" onClick="javascript:location.reload(true)"></td>
					<td width="50">&nbsp;</td>
					<td>
					<%				
						'ObjRS.PageSize	
						If Num > ObjRS.PageSize Then
							abspage = ObjRS.AbsolutePage
							pagecnt = ObjRS.PageCount
							
							Response.Write "<div align=""center"">" & vbcrlf
							Response.Write "<a href="""
							Response.Write Request.ServerVariables("SCRIPT_NAME")
							Response.Write "?pagenum=1""><b>最初のページ</b></a>"
							Response.Write "	|	"								
														
							If abspage = 1 Then
								Response.Write "<span>" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 4 & """>&nbsp;<b>" & abspage + 4 & "</b></a>"
							Elseif abspage = 2 then
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"							
								Response.Write "<span>&nbsp;" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
							Elseif abspage = pagecnt then
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 4 & """>&nbsp;<b>" & abspage - 4 & "</b></a>"															
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"															
								Response.Write "<span>&nbsp;" & abspage & "</span>"
							Elseif abspage = CInt(pagecnt-1) then
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"															
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
								Response.Write "<span>&nbsp;" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"									
							Else
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
								Response.Write "<span>&nbsp;" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"								
							End If
							
							Response.Write "	|	"
							Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
							Response.Write "?pagenum=" & pagecnt & """><b>最後のページ</b></a>"
							Response.Write "</div>" & vbcrlf						
										
						End If
					%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>		
		<td>
			<div id="BDIV1">						
				<% If Num>0 Then%>
				<% If blnSorted Then%>
				<iframe frameborder="0" style="background-color:#00FF00;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400px;height:48px;"></iframe>	
				<% Else%>
				<iframe frameborder="0" style="background-color:#00FF00;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400px;height:38px;"></iframe>	
				<% End If%>			
					<table border="1" cellpadding="0" cellspacing="0" width="100%">										
						<thead>																				
							<th id="H1Col01" class="hlist" nowrap><%=HeaderTbl1(0)%></th>							
							<th id="H1Col02" class="hlist" nowrap><%=HeaderTbl1(1)%></th>
							<th id="H1Col03" class="hlist" nowrap><%=HeaderTbl1(2)%></th>
							<th id="H1Col04" class="hlist" nowrap><%=HeaderTbl1(3)%></th>	
							<th id="H1Col05" class="hlist" nowrap><%=HeaderTbl1(4)%></th>							
							<th id="H1Col06" class="hlist" nowrap><%=HeaderTbl1(5)%></th>
							<th id="H1Col07" class="hlist" nowrap><%=HeaderTbl1(6)%></th>
							<th id="H1Col08" class="hlist" nowrap><%=HeaderTbl1(7)%></th>
							<th id="H1Col09" class="hlist" nowrap><%=HeaderTbl1(8)%></th>
							<th id="H1Col10" class="hlist" nowrap><%=HeaderTbl1(9)%></th>
							<th id="H1Col11" class="hlist" nowrap><%=HeaderTbl1(10)%></th>
							<th id="H1Col12" class="hlist" nowrap><%=HeaderTbl1(11)%></th>
							<th id="H1Col13" class="hlist" nowrap><%=HeaderTbl1(12)%></th>
							<th id="H1Col14" class="hlist" nowrap><%=HeaderTbl1(13)%></th>
						</thead>
						<tbody>							
							<% 
								 For i=1 To ObjRS.PageSize
								 	If Not ObjRS.EOF Then								

							%>								
							<tr class=bgw>													
								<td id="D1Col01" height="22" nowrap><%=Trim(ObjRS("InPutDate"))%><BR></td>	
								<td id="D1Col02" nowrap><%=Trim(ObjRS("SenderCode"))%><BR></td>
								<td id="D1Col03" nowrap><%=Trim(ObjRS("Flag2"))%><BR></td>
								<td id="D1Col04" nowrap>								
								<%
									v_ItemName = "chkAnsNo" + cstr(i)
									Response.Write "<select name= '" & v_ItemName & "' class=chr>"													
									Response.Write "<option value='0'>未</option>"
									Response.Write "<option value='1'>Yes</option>"
									Response.Write "<option value='2'>No</option>"										
									Response.Write "</select>"										
								%>								
								</td>
								<td id="D1Col05" nowrap><%=Trim(ObjRS("BookNo"))%><BR></td>								
								<td id="D1Col06" nowrap><A HREF="JavaScript:GoConinf('<%=Trim(ObjRS("BookNo"))%>');"><%=Trim(ObjRS("NumCount"))%></A><BR></td>								
								<td id="D1Col07" nowrap><%=Trim(ObjRS("ContSize1"))%><BR></td>
								<td id="D1Col08" nowrap><%=Trim(ObjRS("ContType1"))%><BR></td>
								<td id="D1Col09" nowrap><%=Trim(ObjRS("ContHeight1"))%><BR></td>
								<td id="D1Col10" nowrap><%=Trim(ObjRS("ContMaterial1"))%><BR></td>
								<td id="D1Col11" nowrap><%=Trim(ObjRS("ShipLine"))%><BR></td>
								<td id="D1Col12" nowrap><%=Trim(ObjRS("FullName"))%><BR></td>
								<td id="D1Col13" nowrap><%=Trim(ObjRS("TruckerCode"))%><BR></td>								
								<%
								If Trim(ObjRS("SenderCode")) = USER AND Trim(ObjRS("TruckerCode"))<>COMPcd AND Trim(ObjRS("TruckerCode"))<>""  Then
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
								<INPUT type=hidden name="<%=v_ItemName%>" value="<%=Trim(ObjRS("SenderCode"))%>">
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
				<% Else %>
					<table border="1" cellPadding="2" cellSpacing="0">						
					  <TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
					</table>
				<% End If %>			
			 </div>			
		</td>
	</tr>
	<tr>
		<td height="5"></td>
	</tr>
	<tr>		
		<td>	
			<table cellpadding="0" cellspacing="0">
				<td width="20"></td>
				<td width="125" nowrap><B>実搬出入事前情報</B></td>
				<td width="133">&nbsp;</td>
				<td><input type="button" value="表示列設定" onClick="OpenCodeWin('表示列選択（実搬出入）','2')"></td>
				<td width="25">&nbsp;</td>
				<td><input type="button" value="ソート設定"  onClick="OpenCodeWin2('6');"></td>				
				<td width="40">&nbsp;</td>
				<td width="100">&nbsp;</td>
				<td>
					<%					
						If Num2 > 0 Then
						'If Num2 > ObjRS2.PageSize Then
							abspage = ObjRS2.AbsolutePage
							pagecnt = ObjRS2.PageCount
	
							Response.Write "<div align=""center"">" & vbcrlf
							Response.Write "<a href="""
							Response.Write Request.ServerVariables("SCRIPT_NAME")
							Response.Write "?pagenum2=1""><b>最初のページ</b></a>"
							Response.Write "	|	"
							
							If abspage = 1 Then
								Response.Write "<span>" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 4 & """>&nbsp;<b>" & abspage + 4 & "</b></a>"
							Elseif abspage = 2 then
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"							
								Response.Write "<span>&nbsp;" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 3 & """>&nbsp;<b>" & abspage + 3 & "</b></a>"
							Elseif abspage = pagecnt then
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 4 & """>&nbsp;<b>" & abspage - 4 & "</b></a>"															
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"															
								Response.Write "<span>&nbsp;" & abspage & "</span>"
							Elseif abspage = CInt(pagecnt-1) then
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 3 & """>&nbsp;<b>" & abspage - 3 & "</b></a>"															
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
								Response.Write "<span>&nbsp;" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"									
							Else
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 2 & """>&nbsp;<b>" & abspage - 2 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage - 1 & """>&nbsp;<b>" & abspage - 1 & "</b></a>"
								Response.Write "<span>&nbsp;" & abspage & "</span>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 1 & """>&nbsp;<b>" & abspage + 1 & "</b></a>"
								Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
								Response.Write "?pagenum2=" & abspage + 2 & """>&nbsp;<b>" & abspage + 2 & "</b></a>"								
							End If

							Response.Write "	|	"
							Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME")
							Response.Write "?pagenum2=" & pagecnt & """><b>最後のページ</b></a>"
							Response.Write "</div>" & vbcrlf
						End If									
					%>
					</td>	
			</table>			
		</td>
	</tr>
	<tr>
		<td>
			<div id="BDIV2">				
				<% If Num2>0 Then%>
					<% If blnSorted2 Then%>
					<iframe frameborder="0" style="background-color:#00FF00;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400;height:48px;"></iframe>				
					<% Else%>
					<iframe frameborder="0" style="background-color:#00FF00;position:absolute;top:expression(this.offsetParent.scrollTop);left:0px;width:400;height:38px;"></iframe>									
					<% End If%>
					<table border="1" cellpadding="0" cellspacing="0" width="100%">				
						<thead>
							<tr>
								<th id="H2Col01" class="hlist" align="left" nowrap><%=HeaderTbl2(0)%></th>								
								<th id="H2Col02" class="hlist" nowrap><%=HeaderTbl2(1)%></th>
								<th id="H2Col03" class="hlist" nowrap><%=HeaderTbl2(2)%></th>
								<th id="H2Col04" class="hlist" nowrap><%=HeaderTbl2(3)%></th>
								<th id="H2Col05" class="hlist" nowrap><%=HeaderTbl2(4)%></th>									
								<th id="H2Col06" class="hlist" nowrap><%=HeaderTbl2(5)%></th>
								<th id="H2Col07" class="hlist" nowrap><%=HeaderTbl2(6)%></th>
								<th id="H2Col08" class="hlist" nowrap><%=HeaderTbl2(7)%></th>
								<th id="H2Col09" class="hlist" nowrap><%=HeaderTbl2(8)%></th>
								<th id="H2Col10" class="hlist" nowrap><%=HeaderTbl2(9)%></th>
								<th id="H2Col11" class="hlist" nowrap><%=HeaderTbl2(10)%></th>
								<th id="H2Col12" class="hlist" nowrap><%=HeaderTbl2(11)%></th>
								<th id="H2Col13" class="hlist" nowrap><%=HeaderTbl2(12)%></th>
								<th id="H2Col14" class="hlist" nowrap><%=HeaderTbl2(13)%></th>
								<th id="H2Col15" class="hlist" nowrap><%=HeaderTbl2(14)%></th>						
							</tr>
						</thead>																	
						<tbody>
							<% 
								'i = 1
								'While Not ObjRS2.EOF 							
								For i=1 To ObjRS2.PageSize
								 	If Not ObjRS2.EOF Then
							%>
							<% v_ItemName = "chkInOut" + cstr(i) %>
							<% if Trim(ObjRS2("Type")) = "FULLOUT" then%>
								<tr bgcolor="#FFCCFF">
								<td id="D2Col01" align="center" width="50" nowrap><input type="checkbox" name="<%= v_ItemName %>" disabled><BR></td>
							<% else%>
								<tr bgcolor="#CCFFFF">
								<td id="D2Col01" align="center" width="50" nowrap><input type="checkbox" name="<%= v_ItemName %>"><BR></td>
							<% end if%>																	
								
																								
								<td id="D2Col02" nowrap><%=Trim(ObjRS2("WorkDate"))%><BR></td>							
								<td id="D2Col03" nowrap><%=Trim(ObjRS2("Code1"))%><BR></td>
								<td id="D2Col04" nowrap><%=Trim(ObjRS2("Flag2"))%><BR></td>
								<% v_ItemName = "chkAns2No" + cstr(i) %>
								<td id="D2Col05" nowrap>								
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
								<td id="D2Col06" nowrap><A HREF="JavaScript:GoRenew('<%=i%>');"><%=Trim(ObjRS2("WkNo"))%></A><BR></td>																
								<% else %>
								<td id="D2Col06" nowrap><A HREF="JavaScript:GoRenew2('<%=ObjRS2("WkNo")%>','<%=ObjRS2("BookNo")%>','<%=ObjRS2("ContNo")%>');"><%=Trim(ObjRS2("WkNo"))%></A><BR></td>																
								<% end if %>
								<td id="D2Col07" nowrap><A HREF="JavaScript:GoConinf2('<%=Trim(ObjRS2("WkNo"))%>','<%=Trim(ObjRS2("FullOutType"))%>','<%=Trim(ObjRS2("BLContNo"))%>')"><%=Trim(ObjRS2("BLContNo"))%></A><BR></td>																
								<td id="D2Col08" nowrap><%=Trim(ObjRS2("ShipLine"))%><BR></td>
								<td id="D2Col09" nowrap><%=Trim(ObjRS2("ShipName"))%><BR></td>
								<td id="D2Col10" nowrap><%=Trim(ObjRS2("ContSize"))%><BR></td>
								<td id="D2Col11" nowrap><%=Trim(ObjRS2("ReceiveFrom"))%><BR></td>
								<td id="D2Col12" nowrap><%=Trim(ObjRS2("CY"))%><BR></td>
								<% if Trim(ObjRS2("DelPermitDate")) = "-" then%>
									<td id="D2Col13" nowrap align="center"><%=Trim(ObjRS2("DelPermitDate"))%><BR></td>
								<% else%>
									<td id="D2Col13" nowrap><%=Trim(ObjRS2("DelPermitDate"))%><BR></td>
								<% end if%>								
								<td id="D2Col14" nowrap><%=Trim(ObjRS2("FreeTime"))%><BR></td>								
								<td id="D2Col15" nowrap><%=Trim(ObjRS2("CYCut"))%><BR></td>
								
								
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
							</tr>
						<% 
									'i = i + 1
									ObjRS2.MoveNext 		
									End If
								Next	
							'Wend
						
							ObjRS2.close    
						    DisConnDBH ObjConn2, ObjRS2
				
						%>						    									
						</tbody>								
					</table>
				<% Else %>
					<table border="1" cellPadding="2" cellSpacing="0">						
						<TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
					</table>
				<% End If%>
			</div>
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
</div>
</BODY>
</HTML>