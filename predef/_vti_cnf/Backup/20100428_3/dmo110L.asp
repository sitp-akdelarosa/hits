<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo110L.asp				_/
'_/	Function	:空搬入情報一覧画面リスト出力		_/
'_/	Date		:2003/05/27				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-001 2003/07/29	CSV出力対応	_/
'_/			:C-002 2003/07/29	備考欄対応	_/
'_/			:C-004 2003/08/22	表示順整形	_/
'_/			:3th   2004/01/31	3次対応	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
	
	Const CONST_ASC = "<BR><IMG border=0 src=Image/ascending.gif>"
	Const CONST_DESC = "<BR><IMG border=0 src=Image/descending.gif>"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
'CW-055  Session.Contents.Remove("DateP")
'CW-055  Session.Contents.Remove("NumP")

'ユーザデータ所得
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")
'INIファイルより設定値を取得
  dim tmp(2),calcDate1
  getIni tmp
  calcDate1 = DateAdd("d", "-"&tmp(1), Date)

'データ取得
  dim Num, DtTbl,i,j,SortFlag,SortKye

  If Request("SortFlag") = "" Then
    SortFlag = 0
  Else
    SortFlag = Request("SortFlag")
  End If

  'ソートケース
  dim strWrer,ErrerM

'2009/02/25 Add-S G.Ariola   
    dim strOrder
  dim FieldName
   ReDim FieldName(13)
  
  'FieldName=Array("ITC.WorkDate","Code1","Name1","ITC.ContNo","mV.ShipLine"," mV.FullName"," CNT.ContSize","INC.ReturnPlace","INC.DetentionFreeTime","Code2","Flag1","ITC.Comment1")
  FieldName=Array("WorkDate","Code1","ContNo","ShipLine"," FullName"," ContSize","ReturnPlace","DetentionFreeTime","Code2","Flag1","Comment1","Name1")
  
  strOrder = getSort(Session("Key1"),Session("KeySort1"),"")
  strOrder = getSort(Session("Key2"),Session("KeySort2"),strOrder)
  strOrder = getSort(Session("Key3"),Session("KeySort3"),strOrder)
'2009/02/25 Add-E G.Ariola
  
  Select Case SortFlag
      Case "0" '初期表示:搬入予定日順に表示
        WriteLogH "b201", "空搬入事前情報一覧", "01", ""
        '2010/04/23 M.Marquez Upd-S
        'strWrer="AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
        strWrer="AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
        '2010/04/23 M.Marquez Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 0
      Case "1" '指示先が未照会のコンテナ一覧
        WriteLogH "b201", "空搬入事前情報一覧", "03", ""
        '2010/04/23 M.Marquez Upd-S
        'strWrer="AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
        strWrer="AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
        '2010/04/23 M.Marquez Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 1
'3th          j=1
'3th          DtTbl(0)(8)=0
'3th          For i=1 To Num
'3th            If DtTbl(i)(6) = "未" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(j)(7)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "7" '保留
        WriteLogH "b201", "空搬入事前情報一覧", "07", ""
        '2010/04/23 M.Marquez Upd-S
        'strWrer="AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
        strWrer="AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
        '2010/04/23 M.Marquez Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 2
'3th          j=1
'3th          DtTbl(0)(8)=0
'3th          For i=1 To Num
'3th            If DtTbl(i)(6) = "No" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(j)(7)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "2" '搬入未完了分をすべて表示
        WriteLogH "b201", "空搬入事前情報一覧", "02", ""
        strWrer = "AND ITC.WorkCompleteDate IS Null "
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 0
      Case "3" '全件表示
        WriteLogH "b201", "空搬入事前情報一覧", "04", ""
        strWrer = " "
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 0

      Case "4" 'コンテナ番号で検索
          SortKye=Request("SortKye")
          WriteLogH "b201", "空搬入事前情報一覧","11",SortKye
'CW-055 Chenge Start
'          If Session.Contents("ConNum") = "" Then
'            jampErrerP "0","b201","11","空搬入：一覧検索(コンテナ番号)","001",""
'          Else
'            DtTbl=Session("DateT")
'            Num  =Session.Contents("ConNum")
'          End If
'3th chage          Get_Data Num,DtTbl
          strWrer = "AND ITC.ContNo LIKE '%" & SortKye & "'"
          getData DtTbl,strWrer,0
'CW-055 Chenge End
'3th          j=1
'3th          DtTbl(0)(8)=0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(3),Len(SortKye))= SortKye Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(j)(7)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "8" '照会済
          WriteLogH "b207", "空搬入事前情報照会済","01",SortKye
'CW-055 Chenge Start
'          If Session.Contents("ConNum") = "" Then
'            jampErrerP "0","b207","01","空搬入：一覧照会済","001",""
'          Else
'            DtTbl=Session("DateT")
'            Num  =Session.Contents("ConNum")
'          End If
          Get_Data Num,DtTbl
'CW-055 Chenge End
        'エラートラップ開始
          on error resume next
        'DB接続
          dim ObjConn, ObjRS, StrSQL
          ConnDBH ObjConn, ObjRS
          For i=1 To Num
'CW-002            If DtTbl(i)(7) <> 0 Then
            If DtTbl(i)(7) <> 0 AND DtTbl(i)(4)="" AND DtTbl(i)(8)="未" Then
              StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                       "UpdtTmnl='"& USER &"', TruckerFlag"& DtTbl(i)(7) &"=1 "&_
                       "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                       "WHERE ContNo='"& DtTbl(i)(3) &"' AND WkType='2'  AND Process='R')"
              ObjConn.Execute(StrSQL)
              if err <> 0 then
                Set ObjRS = Nothing
                jampErrerPDB ObjConn,"0","b207","01","空搬入：一覧照会済","103","SQL:<BR>"&strSQL
              end if
            End If
          Next
        'DB接続解除
          DisConnDBH ObjConn, ObjRS
        'エラートラップ解除
          on error goto 0
          Response.Redirect "./dmo110L.asp"
	 Case else '全件表示
          WriteLogH "b101", "実搬出事前情報一覧", "04",""
          strWrer = " "
          getData DtTbl,strWrer,0
  End Select
'CW-055  Session.Contents.Remove("DateT")
'CW-055  Session("DateT")=DtTbl
'CW-055  Session.Contents("ConNum")=Num
'CW-055  If Num=0 Then
'CW-055    Session.Contents("NullFlag")=0
'CW-055  Else
'CW-055    Session.Contents("NullFlag")=1
'CW-055  End If

'データ取得関数
'2009/02/25 Add-S G.Ariola
Function getSort(Key,SortKey,str)
getSort = str
	if Key <> "" then
	
		if str = "" then
			'getSort = " ORDER BY (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
			if (FieldName(Key) = "WorkDate" OR FieldName(Key) = "DetentionFreeTime") AND SortKey = "ASC" then 
			getSort = " ORDER BY isnull(" & FieldName(Key) & ",DATEADD(Year,100,getdate())) " & SortKey	
			else
			getSort = " ORDER BY " & FieldName(Key) & " " & SortKey	
			end if
			
		else
			'getSort = str & " , (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
			if (FieldName(Key) = "WorkDate" OR FieldName(Key) = "DetentionFreeTime") AND SortKey = "ASC"  then 
			getSort = str & " , isnull(" & FieldName(Key) & ",DATEADD(Year,100,getdate())) " & SortKey	
			else
			getSort = str & " , " & FieldName(Key) & " " & SortKey	
			end if
			
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
'2009/02/25 Add-E G.Ariola
 
'3thFunction getData(DtTbl,strWrer)
Function getData(DtTbl,strWrer,DelType)
	ReDim DtTbl(1)
'C-002  DtTbl(0)=Array("入力日","搬入予定日","指示元","コンテナ番号","完了日時","指示先","指示先<BR>回答","照会先Frag","指示元への回答","船社","船名","サイズ","返却先","ディテンション<BR>フリータイム")
'20030911  DtTbl(0)=Array("入力日","搬入予定日","指示元","コンテナ番号","完了日時","指示先","指示先<BR>回答","照会先Frag","指示元への回答","船社","船名","サイズ","返却先","ディテンション<BR>フリータイム","備考")
	DtTbl(0)=Array("入力日","搬入<BR>予定日","指示元","コンテナ番号","完了日時","指示先","指示先<BR>回答","照会先Frag","指示元への回答","船社","船名","SZ","返却先","ディテンション<BR>フリータイム","備考","作業管理番号","コード","指示元担当")
	dim ctr
	for ctr = 1 to 3
		Session(CSTR("Key" & ctr))
		if Session(CSTR("Key" & ctr)) <> "" then
			Select Case Session(CSTR("Key" & ctr))
				Case "0" '搬入予定日
					DtTbl(0)(1) = DtTbl(0)(1) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "1" '指示元 － コード
					DtTbl(0)(16) = DtTbl(0)(16) & getImage(Session(CSTR("KeySort" & ctr)))				
				Case "2" 'コンテナ番号
					DtTbl(0)(3) = DtTbl(0)(3) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "3" '船社
					DtTbl(0)(9) = DtTbl(0)(9) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "4" '船名
					DtTbl(0)(10) = DtTbl(0)(10) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "5" 'SZ
					DtTbl(0)(11) = DtTbl(0)(11) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "6" '返却先
					DtTbl(0)(12) = DtTbl(0)(12) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "7" 'ディテンションフリータイム
					DtTbl(0)(13) = DtTbl(0)(13) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "8" '指示先 － コード
					DtTbl(0)(5) = DtTbl(0)(5) & getImage(Session(CSTR("KeySort" & ctr)))
'				Case "9" '指示先 － 担当
'					DtTbl(0)(19) = DtTbl(0)(19) & getImage(Session(CSTR("KeySort" & ctr)))		
				Case "9" '指示先回答
					DtTbl(0)(6) = DtTbl(0)(6) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "10" '備考
					DtTbl(0)(14) = DtTbl(0)(14) & getImage(Session(CSTR("KeySort" & ctr)))
				Case "11" '指示元 － 担当
					DtTbl(0)(17) = DtTbl(0)(17) & getImage(Session(CSTR("KeySort" & ctr)))
			End Select
		end if
	next
'2009/02/25 Add-E G.Ariola

	DtTbl(0)(8) =0
'3th Add Start
	Dim DelStr,DelTarget
	DelStr=Array("","未","No")
	DelTarget=Array(0,6,6)
'3th Add End

	'エラートラップ開始
	on error resume next
	'DB接続
	dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	'対象件数取得
	StrSQL = "SELECT count(WkContrlNo) AS CNUM FROM hITCommonInfo ITC "&_
		"WHERE WkType='2' AND (RegisterCode='"& USER &"' "&_
		"OR TruckerSubCode1='"& COMPcd &"' OR TruckerSubCode2='"&_
		COMPcd &"' OR TruckerSubCode3='"& COMPcd &"' OR TruckerSubCode4='"& COMPcd &"') AND Process='R' " &_
		strWrer 

	ObjRS.Open StrSQL, ObjConn

	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jampErrerP "2","b201","01","空搬入：一覧","101","SQL:<BR>"&StrSQL
		Exit Function
	end if

	Num = ObjRS("CNUM")
	ObjRS.close
	ReDim Preserve DtTbl(Num)

	'データ取得
	StrSQL = "SELECT T.* FROM (SELECT ITC.InputDate, ITC.WorkDate, ITC.RegisterCode, ITC.TruckerSubCode1, ITC.TruckerSubCode2, "&_
		"ITC.TruckerSubCode3, ITC.TruckerSubCode4,ITC.ContNo, ITC.WorkCompleteDate, "&_
			 "(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode1 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			"ELSE CASE WHEN mU.UserType='5' THEN mU.HeadCompanyCode ELSE ITC.RegisterCode END "&_
			"END) as Code1, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName1 "&_
			"ELSE ITC.TruckerSubName1 "&_
			"END) as Name1, "&_
			"(CASE WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN '' "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubCode4 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubCode3 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubCode2 "&_
			"ELSE ITC.TruckerSubCode1 "&_
			"END) as Code2, "&_
			"(CASE "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITC.TruckerSubName2 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITC.TruckerSubName3 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITC.TruckerSubName4 "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"ELSE ITC.TruckerSubName1 END) as Name2, "&_
			"(CASE "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN NULL "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag2 "&_								
			"ELSE ITR.TruckerFlag1 END) as Flag1, "&_			 
			"(CASE "&_
			"WHEN ITC.TruckerSubCode4 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode3 = '"& COMPcd &"' THEN ITR.TruckerFlag3 "&_
			"WHEN ITC.TruckerSubCode2 = '"& COMPcd &"' THEN ITR.TruckerFlag4 "&_
			"WHEN ITC.TruckerSubCode1 = '"& COMPcd &"' THEN ITR.TruckerFlag1 "&_								
			"ELSE NULL END) as Flag2, "&_			 
		"ITC.WkContrlNo, ITC.Comment1, ITR.TruckerFlag1, "&_
		"ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, mV.ShipLine, mV.FullName, CNT.ContSize, "&_
		"INC.ReturnPlace, INC.DetentionFreeTime, mU.HeadCompanyCode, mU.UserType "
		
	  StrSQL = StrSQL &	"FROM ((((hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo) "&_
		"INNER JOIN Container AS CNT ON ITC.ContNo = CNT.ContNo) "&_
		"LEFT JOIN mVessel AS mV ON CNT.VslCode = mV.VslCode) "&_
		"INNER JOIN ImportCont AS INC ON (CNT.ContNo = INC.ContNo) AND (CNT.VoyCtrl = INC.VoyCtrl) "&_
		"AND (CNT.VslCode = INC.VslCode))"&_
		"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
		"WHERE WkType='2' AND (RegisterCode='"& USER &"' "&_
		"OR TruckerSubCode1='"& COMPcd &"' OR TruckerSubCode2='"&_
		COMPcd &"' OR TruckerSubCode3='"& COMPcd &"' OR TruckerSubCode4='"& COMPcd &"') AND Process='R' "& strWrer &") AS T "&_
      strOrder
	

'             "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
'CW-051 ADD This Line "mU.HeadCompanyCode, mU.UserType "&_
'CW-051 ADD This Line "INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
'C-004 Chenge This Line "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC ASC"
'20030911 ADD This Item "ITC.WkContrlNo, "
'response.Write(StrSQL)
'response.End()
'C-002 ADD 

	ObjRS.Open StrSQL, ObjConn

	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jampErrerP "2","b201","01","空搬入：一覧","102","SQL:<BR>"&StrSQL
		Exit Function
	end if

	i=1

	Do Until ObjRS.EOF
'ADD 20050530 for CSV出力に出力されないレコードが存在する不具合周辺確認 重複時の排他
		If DtTbl(i-1)(3) <> Trim(ObjRS("ContNo")) Then
'C-002      DtTbl(i)=Array("","","","","","","","","","","","","","")
'20030911      DtTbl(i)=Array("","","","","","","","","","","","","","","")
        'DtTbl(i)=Array("","","","","","","","","","","","","","","","")
			DtTbl(i)=Array("","","","","","","","","","","","","","","","","","")
			DtTbl(i)(0)=Mid(ObjRS("InPutDate"),3,8)
			DtTbl(i)(1)=Mid(ObjRS("WorkDate"),3,8)
			DtTbl(i)(3)=Trim(ObjRS("ContNo"))
			DtTbl(i)(4)=Trim(Mid(ObjRS("WorkCompleteDate"),3,8))
			DtTbl(i)(9)=Trim(ObjRS("ShipLine"))
'C-001      DtTbl(i)(10)=Left(ObjRS("FullName"),12)
			DtTbl(i)(10)=Trim(ObjRS("FullName"))
			DtTbl(i)(11)=Trim(ObjRS("ContSize"))
			DtTbl(i)(12)=Trim(ObjRS("ReturnPlace"))
			DtTbl(i)(13)=Trim(ObjRS("DetentionFreeTime"))
			DtTbl(i)(14)=Trim(ObjRS("Comment1"))		'C-002
			DtTbl(i)(15)=Trim(ObjRS("WkContrlNo"))		'20030911
		
'2009/02/25 Add-S G.Ariola		
			DtTbl(i)(2) = Trim(ObjRS("Code1"))
			DtTbl(i)(16) = Trim(ObjRS("Name1"))
			DtTbl(i)(5) = Trim(ObjRS("Code2"))
			DtTbl(i)(17) = Trim(ObjRS("Name2"))

			DtTbl(i)(6) = ObjRS("Flag1")
			DtTbl(i)(8) = ObjRS("Flag2")

			If Trim(ObjRS("TruckerSubCode4")) = COMPcd Then
				DtTbl(i)(7) = 4
			ElseIf Trim(ObjRS("TruckerSubCode3")) = COMPcd Then
				DtTbl(i)(7) = 3
			ElseIf Trim(ObjRS("TruckerSubCode2")) = COMPcd Then
				DtTbl(i)(7) = 2
			ElseIf Trim(ObjRS("TruckerSubCode1")) = COMPcd Then
				DtTbl(i)(7) = 1
			Else
				DtTbl(i)(7) = 0
			end if
'2009/02/25 Add-E G.Ariola	  
		
      '指示先照会済みフラグ
'        If Trim(ObjRS("TruckerSubCode4")) = COMPcd Then
'          DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode3"))
'          DtTbl(i)(5) = Null
'          DtTbl(i)(7) = 4
'          DtTbl(i)(6) = Null
'          DtTbl(i)(8) = ObjRS("TruckerFlag4")
'        ElseIf Trim(ObjRS("TruckerSubCode3")) = COMPcd Then
'          DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode2"))
'          DtTbl(i)(5) = Trim(ObjRS("TruckerSubCode4"))
'          DtTbl(i)(7) = 3
'          DtTbl(i)(6) = ObjRS("TruckerFlag4")
'          DtTbl(i)(8) = ObjRS("TruckerFlag3")
'        ElseIf Trim(ObjRS("TruckerSubCode2")) = COMPcd Then
'          DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode1"))
'          DtTbl(i)(5) = Trim(ObjRS("TruckerSubCode3"))
'          DtTbl(i)(7) = 2
'          DtTbl(i)(6) = ObjRS("TruckerFlag3")
'          DtTbl(i)(8) = ObjRS("TruckerFlag2")
'        ElseIf Trim(ObjRS("TruckerSubCode1")) = COMPcd Then
'          If ObjRS("UserType") = "5" Then           'CW-051
'            DtTbl(i)(2) = Trim(ObjRS("HeadCompanyCode")) 'CW-051
'          Else                      'CW-051
'            DtTbl(i)(2) = Trim(ObjRS("RegisterCode"))
'          End If                        'CW-051
'          DtTbl(i)(5) = Trim(ObjRS("TruckerSubCode2"))
'          DtTbl(i)(7) = 1
'          DtTbl(i)(6) = ObjRS("TruckerFlag2")
'          DtTbl(i)(8) = ObjRS("TruckerFlag1")
'        Else
'          If ObjRS("UserType") = "5" Then           'CW-051
'            DtTbl(i)(2) = Trim(ObjRS("HeadCompanyCode")) 'CW-051
'          Else                      'CW-051
'            DtTbl(i)(2) = Trim(ObjRS("RegisterCode"))
'          End If                        'CW-051
'          DtTbl(i)(5) = Trim(ObjRS("TruckerSubCode1"))
'          DtTbl(i)(7) = 0
'          DtTbl(i)(6) = ObjRS("TruckerFlag1")
'          DtTbl(i)(8) = Null
'        End If

		
			If IsNull(DtTbl(i)(5)) Then
				DtTbl(i)(6) ="　"
			ElseIf DtTbl(i)(6) = 0 Then
				DtTbl(i)(6) ="未"
			ElseIf DtTbl(i)(6) = 1 Then
				DtTbl(i)(6) ="Yes"
			Else
				DtTbl(i)(6) ="No"
			End If

			If DtTbl(i)(8)=0 Then
				DtTbl(i)(8) ="未"
			ElseIf DtTbl(i)(8) = 1 Then
				DtTbl(i)(8) ="Yes"
			ElseIf DtTbl(i)(8) = 2 Then
				DtTbl(i)(8) ="No"
			Else
				DtTbl(i)(8) ="　"
			End If
'3th Add Start
			If DelType=0 OR DtTbl(i)(DelTarget(DelType)) = DelStr(DelType) Then
				DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(i)(7)
				i=i+1
			Else
				Num=Num-1
			End If
'      i=i+1
'3th Add End
		End If
		ObjRS.MoveNext   'ADD 20050530 
	Loop
	ObjRS.close
	If i-1<Num Then
		ErrerM = "<DIV class=alert>登録データのうち"& Num-i+1 &"件について関連データ取得失敗のため"&_
				"表示されていません。<BR>システム管理者に問い合わせてください。</DIV><P>"
		Num=i-1
	End If
	
	'DB接続解除
	DisConnDBH ObjConn, ObjRS
	'エラートラップ解除
	on error goto 0
End Function

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空搬入情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//データが無い場合の表示制御
function vew(){
<%If Num<>0 Then%> 
 //var obj1=document.getElementById("HDIV");
 //var obj2=document.getElementById("VDIV");
 var obj3=document.getElementById("BDIV");
//alert((document.body.offsetWidth-10))
/* if((document.body.offsetWidth-10)  < 877){
    obj1.style.width=document.body.offsetWidth-10;
	obj1.style.overflow="auto";
	}
 else
 {
 	 obj1.style.width=document.body.offsetWidth-10;
	 obj1.style.overflowX="auto";
 }
 obj2.style.height=document.body.offsetHeight-80;*/
 //alert(document.body.offsetWidth);
 if((document.body.offsetWidth-10)  < 802){
    obj3.style.width=document.body.offsetWidth;
    obj3.style.overflowX="auto";
 }
 else{
  	 obj3.style.width=document.body.offsetWidth-10;
  	 obj3.style.overflowX="auto";
 }
 obj3.style.height=document.body.offsetHeight;
 obj3.style.overflowY="auto";
<% End If %>
}
//更新
function GoRenew(conNo,wkconNo){
  Fname=document.dmo110F;
  Fname.CONnum.value=conNo;
  Fname.WkconNo.value=wkconNo;
  Fname.action="./dmi115.asp";
// 2009/10/09 Upd-S Fujiyama
//  newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes");
  newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes,scrollbars=yes");
// 2009/10/09 Upd-E Fujiyama
  Fname.target="ReEntry";
  Fname.submit();
}
//検索
function SerchC(SortFlag,Kye){
  Fname=document.dmo110F;
  Fname.SortFlag.value=SortFlag;
  Fname.SortKye.value=Kye;
  Fname.target="_self";
  Fname.action="./dmo110L.asp";
  Fname.submit();
}
//照会済
function GoSyokaizumi(){
  target=document.dmo110F;
  if(target.DataNum.value>0){
    flag = confirm('未回答の回答を「Yes」にしますか？');
    if(flag==true){
      len=target.elements.length;
      for(i=0;i<len;i++){
        target.elements[i].disabled=false;
      }
      target.SortFlag.value=8;
      target.target="_self";
      target.action="./dmo110L.asp";
      target.submit();
    }
  }
}
//CSV		ADD C-001
function GoCSV(){
  target=document.dmo110F;
  len=target.elements.length;
  for(i=0;i<len;i++){
    target.elements[i].disabled=false;
  }
  target.target="Bottom";
  target.action="./dmo180.asp";
  target.submit();
}
function showContent(){
    var target=null;
    while (target==null) {
	    target=parent.window.frames(0);
	}
    var target1 = target.window.document.getElementById("loading");
    target1.style.display='none';
    //show content
    document.getElementById("content").style.display='block';
}
// -->
</SCRIPT>
<style type="text/css">
INPUT.chrReadOnly
{
    BORDER-BOTTOM: 0px solid;
    BORDER-LEFT: 0px solid;
    BORDER-RIGHT: 0px solid;
    BORDER-TOP: 0px solid;
	PADDING-BOTTOM: 0px solid;
    PADDING-LEFT: 0px solid;
    PADDING-RIGHT: 0px solid;
    PADDING-TOP: 0px solid;
	margin-bottom:0px solid;
	margin-left:0px solid;
	margin-right:0px solid;
	margin-top:0px solid;
    FONT-FAMILY: 'ＭＳ Ｐゴシック';
    FONT-SIZE: 10pt;
    TEXT-ALIGN: left
}
DIV.BDIV
{
    position: relative;
    border-width: 0px 0px 1px 0px;
}
thead tr 
{
    position: relative;
    top: expression(this.offsetParent.scrollTop);
}
th.hlist 
{
    position: relative;
}
table {
    border-width: 0px 1px 1px 0px;
}
th {
    border-width: 1px 1px 1px 1px;
    padding: 4px;
    background-color: #ffcc33;
}
</style>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="vew();" onResize="vew()" onBlur="vew()">
<!--setTimeout('showContent()', 500); -->
<!-------------空搬入情報一覧画面List--------------------------->
<!--<div id="content" style="display:none;"> -->
<%=ErrerM%>
<Form name="dmo110F" method="POST">
<!--DIV id="HDIV" style=" overflow-x:scroll;">
<DIV style=" width:877; height:41;"--><!--2009/02/25 G.Ariola -->
<div id="BDIV">
<TABLE border="1" cellPadding="2" cellSpacing="0" >
    <%If Num<>0 Then%> 
        <% If DtTbl(0)(8)=0 Then %>
        <thead>
            <tr >
                <th class="hlist" nowrap><%=DtTbl(0)(1)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(2)%><BR><%=DtTbl(0)(16)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(3)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(9)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(10)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(11)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(12)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(13)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(5)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(6)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(14)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(17)%>
                <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>'>
                </th>
            </tr>
            <!--2009/02/25 Add-S G.Ariola test1 -->
            <!--<tr >
	            <th width="100" nowrap><'%=DtTbl(0)(16)%></th>
	            <th width="100" nowrap><'%=DtTbl(0)(17)%></th> -->
	            <!--<TH width="60"><%'=DtTbl(0)(18)%></TH>
	            <TH width="60"><%'=DtTbl(0)(19)%></TH> -->
	            <!--/TH-->
            <!--</tr> -->
	        <!--2009/02/25 Add-E G.Ariola -->
            <!--/TABLE--> <!--2009/02/25 G.Ariola -->
            <!--/DIV--><!--2009/02/25 G.Ariola -->
       </thead>

        <%'If Num>10 Then%><!--DIV id="VDIV" style=" width:848; height:242; overflow-y:auto;"--><!--2009/02/25 G.Ariola -->
        <%'else%><!--DIV id="VDIV" style=" width:832; height:242;"--><!--2009/02/25 G.Ariola -->
        <%'end if%> <!--2009/02/25 G.Ariola -->

        <!--TABLE border="1" cellPadding="2" cellSpacing="0" cols="<%=Num+20%>"--><!--2009/02/25 G.Ariola -->    
       <tbody>
        <% For j=1 to Num %>
          <TR class=bgw>
            <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
            <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(3)%>','<%=DtTbl(j)(15)%>');"><%=DtTbl(j)(3)%></A><BR></TD>
            <TD nowrap><%=DtTbl(j)(9)%><BR></TD>
            <TD nowrap><%=Left(DtTbl(j)(10),12)%><BR></TD>
            <TD nowrap><%=DtTbl(j)(11)%><BR></TD>
            <TD nowrap><%=DtTbl(j)(12)%><BR></TD>
            <TD nowrap><%=DtTbl(j)(13)%><BR></TD>
            <TD nowrap><%=DtTbl(j)(5)%><BR></TD><!--<TD width="57"><INPUT value="<%=DtTbl(j)(17)%>" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 ><BR></TD> -->
            <TD nowrap><%=DtTbl(j)(6)%><BR></TD>
            <TD nowrap><%=Left(DtTbl(j)(14),10)%><BR></TD>
            <TD nowrap><%=Left(DtTbl(j)(16),8)%><!--<INPUT value="" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 > --><BR>
            <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>'>
            </TD>
          </TR>
        <% Next %>
        <!--/TABLE-->	<!--2009/02/25 G.Ariola -->
      </tbody>
        <%'If Num>10 Then%><!--/DIV--><%'end if%>	<!--2009/02/25 G.Ariola -->		
        <% Else %>
            <thead>  
              <tr >
                <th class="hlist" nowrap><%=DtTbl(0)(1)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(2)%><BR><%=DtTbl(0)(16)%></th>
                <!--<th class="hlist" nowrap>指示元<BR>への回答</th> -->
                <th class="hlist" nowrap><%=DtTbl(0)(3)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(9)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(10)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(11)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(12)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(13)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(5)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(6)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(14)%></th>
                <th class="hlist" nowrap><%=DtTbl(0)(17)%>
                <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>'>
                </th>
              </tr>
              <!--2009/02/25 Add-S G.Ariola -->
              <!--<tr >
	            <th width="100"><'%=DtTbl(0)(16)%></th>
	            <th width="100"><'%=DtTbl(0)(17)%></th> -->
	            <!--<TH width="60"><%'=DtTbl(0)(18)%></TH>
	            <TH width="60"><%'=DtTbl(0)(19)%></TH> -->
	           <!--</tr> -->
              <!--2009/02/25 Add-S G.Ariola -->
            </thead>  
             <!--/TABLE--> <!--2009/02/25 G.Ariola -->
             <!--/DIV--><!--2009/02/25 G.Ariola -->

            <%'If Num>10 Then%>  <!--DIV id="VDIV" style=" width:848; height:242; overflow-y:auto;"--><!--2009/02/25 G.Ariola -->
            <%'else%><!--DIV id="VDIV" style=" width:832; height:242;"--><!--2009/02/25 G.Ariola -->
            <%'end if%> <!--2009/02/25 G.Ariola -->
            <tbody>
                <!--TABLE border="1" cellPadding="2" cellSpacing="0" cols="<%=Num+20%>"-->    
                <% For j=1 to Num %>
                  <TR class=bgw>
		            <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
                    <!--<TD nowrap><'%=DtTbl(j)(8)%><BR></TD>  -->
                    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(3)%>','<%=DtTbl(j)(15)%>');"><%=DtTbl(j)(3)%></A><BR></TD>
                    <TD nowrap><%=DtTbl(j)(9)%><BR></TD>
                    <TD nowrap><%=Left(DtTbl(j)(10),12)%><BR></TD>
                    <TD nowrap><%=DtTbl(j)(11)%><BR></TD>
                    <TD nowrap><%=DtTbl(j)(12)%><BR></TD>
                    <TD nowrap><%=DtTbl(j)(13)%><BR></TD>
                    <TD nowrap><%=DtTbl(j)(5)%><BR></TD><!--<TD width="57"><INPUT value="<%=DtTbl(j)(17)%>" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 ><BR></TD> -->
                    <TD nowrap><%=DtTbl(j)(6)%><BR></TD>
                    <TD nowrap><%=Left(DtTbl(j)(14),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(16),8)%><!--<INPUT value="" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 > --><BR>
                    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>'>
                    </TD>
                  </TR>
                <% Next %>
                <!--/TABLE-->	<!--2009/02/25 G.Ariola -->
            </tbody>  
            <%'If Num>10 Then%>  <!--/DIV--><%'end if%>	<!--2009/02/25 G.Ariola -->	
        <% End If %>
        <!--/DIV-->
    <% Else %>
        <TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
    <% End If %>
</TABLE>
</div>
<%'3th del Set_Data Num,DtTbl %>
  <INPUT type=hidden name=DataNum value="<%=Num%>">
  <INPUT type=hidden name=SortFlag value="<%=SortFlag%>" >
  <INPUT type=hidden name=SortKye value="<%=SortKye%>" >
  <INPUT type=hidden name=InfoFlag value="0">
  <INPUT type=hidden name=CONnum value="" >
  <INPUT type=hidden name=WkconNo value="" >
</Form>
<!--</div> -->
<!-------------画面終わり--------------------------->
</BODY></HTML>