<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo010L.asp				_/
'_/	Function	:実搬出情報一覧画面リスト出力		_/
'_/	Date		:2003/05/27				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-001 2003/07/29	CSV出力対応	_/
'_/			:C-002 2003/07/29	備考欄対応	_/
'_/			:C-003 2003/08/22	作業番号での検索_/
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

'ユーザデータ所得
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")
'INIファイルより設定値を取得
  dim param(2),calcDate1
  getIni param
  calcDate1 = DateAdd("d", "-"&param(1), Date)
'データ取得
  dim Num, DtTbl,i,j,SortFlag,SortKye,InfoFlag,Siji
  dim Num2
  dim ObjConn, ObjRS
  dim RecCtr, abspage ,pagecnt
  const gcPage = 10
  
  Siji  =Array("","指定あり","指定なし","一覧","ＢＬ")

  If Request("SortFlag") = "" Then
    SortFlag = 0
  Else
    SortFlag = Request("SortFlag")
    '2010/11/10 M.Marquez Upd-S
    'if Instr(1,SortFlag,",") > 0 then 
    '    SortFlag = Mid(SortFlag,1,Instr(1,SortFlag,",")-1)
    'end if
    '2010/11/10 M.Marquez Upd-E    
  End If

  If Request("InfoFlag") = "" Then
    InfoFlag = 0
  Else
    InfoFlag = Request("InfoFlag")
  End If

  'ソートケース
  dim strWrer,ErrerM

'2009/02/25 Add-S G.Ariola   
    dim strOrder
  dim FieldName
   ReDim FieldName(19)
  
'  FieldName=Array("ISNULL(ITC.WorkDate,DATEADD(Year,100,getdate()))","Code1","Name1","ITC.WkNo","ITC.FullOutType","BLContNo","mV.ShipLine"," ShipName","CON.ContSize","CY","INC.FreeTime","ITC.DeliverTo1","ISNULL(ITC.WorkCompleteDate,DATEADD(Year,100,getdate()))","ITC.ReturnDateStr","ReturnValue","Code2","Name2","Flag1","ITC.Comment1","ITC.Comment2")
  FieldName=Array("WorkDate","Code1","WkNo","FullOutType","BLContNo","ShipLine","ShipName","ContSize","CY","FreeTime","DeliverTo1","WorkCompleteDate","ReturnDateStr","ReturnValue","Code2","Flag1","Comment1","Comment2","Name1")
  
  strOrder = getSort(Session("Key1"),Session("KeySort1"),"")
  strOrder = getSort(Session("Key2"),Session("KeySort2"),strOrder)
  strOrder = getSort(Session("Key3"),Session("KeySort3"),strOrder)
'2009/02/25 Add-E G.Ariola

  Select Case SortFlag
  
'20030910 初期表示を搬出予定日順に表示(当日以降のみ)に変更
      Case "0" '初期表示:搬出予定日順に表示(当日以降のみ)
          WriteLogH "b101", "実搬出事前情報一覧", "01", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "&_
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "&_
		   	 	    "AND ('"&Date&"'<=ITC.WorkDate Or ITC.WorkDate IS Null) "
                    '"AND (DateDiff(day,ITC.WorkDate,'"&Date&"')<=0 Or ITC.WorkDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
''3th          getData DtTbl,strWrer
          '2009/11/02 Add-S Tanaka
          'strOrder=" ORDER BY isnull(T.WorkDate,DATEADD(Year,100,getdate())),T.InputDate ASC "
          '2009/11/02 Add-E Tanaka
          GetData DtTbl, strWrer, 1

''3th         j=1
''3th         DtTbl(0)(14) = 0
''3th          For i=1 To Num
''3th            If DtTbl(i)(8)  <> "済" Then
''3th              DtTbl(j)=DtTbl(i)
''3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
''3th              j=j+1
''3th            End If
''3th          Next
''3th          Num=j-1
      Case "12" '旧初期表示:搬出予定日順に表示(未完了分も表示)
          WriteLogH "b101", "実搬出事前情報一覧", "01", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 1
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If DtTbl(i)(8)  <> "済" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
'Del 030722      Case "1" '返却を要するコンテナ順
      Case "2" '未照会
          WriteLogH "b101", "実搬出事前情報一覧", "03", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 2
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If DtTbl(i)(10) = "未" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "7" '保留
          WriteLogH "b101", "実搬出事前情報一覧", "07", ""
		  '2010/04/23 G.Ariola Upd-S
          'strWrer = "AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
		  strWrer = "AND ('"&calcDate1&"' <= ITC.WorkCompleteDate Or ITC.WorkCompleteDate IS Null) "
		  '2010/04/23 G.Ariola Upd-E
'3th          getData DtTbl,strWrer
          GetData DtTbl, strWrer, 3
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th'CW-031            If DtTbl(i)(10) = "no" Then
'3th            If DtTbl(i)(10) = "No" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "3" 'コンテナ番号で検索
          SortKye=Request("SortKye")
          WriteLogH "b101", "実搬出事前情報一覧", "11",SortKye
'3th chage          Get_Data Num,DtTbl
          strWrer = "AND ITC.ContNo LIKE '%" & SortKye & "'"
          getData DtTbl,strWrer,0

'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(5),Len(SortKye))= SortKye AND DtTbl(i)(4) <> 4 Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
'2010/04/24 Upd-S Tanaka BL番号での検索が全件となっているので変更(旧ソースに戻す)
'      Case "4" 'コンテナ番号で検索
'          SortKye=Request("SortKye")
'          WriteLogH "b101", "実搬出事前情報一覧", "11",SortKye
''3th chage          Get_Data Num,DtTbl
'          strWrer = "AND ITC.ContNo LIKE '%" & SortKye & "'"
'          getData DtTbl,strWrer,0

      Case "4" 'BL番号で検索
          SortKye=Request("SortKye")
          WriteLogH "b101", "実搬出事前情報一覧", "11",SortKye
'3th chage          Get_Data Num,DtTbl
          strWrer = "AND ITC.BLNo LIKE '%" & SortKye & "' AND ITC.FullOutType='4'"
          getData DtTbl,strWrer,0


'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(5),Len(SortKye))= SortKye AND DtTbl(i)(4) <> 4 Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "11" '作業番号で検索
          SortKye=Request("SortKye")
          WriteLogH "b101", "実搬出事前情報一覧", "11",SortKye
'3th          Get_Data Num,DtTbl
          strWrer = "AND ITC.WkNo LIKE '%" & SortKye & "'"
          getData DtTbl,strWrer,0
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(3),Len(SortKye))= SortKye Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
'ADD End C-003
      Case "5" '全件表示
          WriteLogH "b101", "実搬出事前情報一覧", "04",""
          strWrer = " "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,0
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(i)(13)
'3th          Next
      Case "6" '搬出未完了分をすべて表示
          WriteLogH "b101", "実搬出事前情報一覧", "06",""
          strWrer = "AND ITC.WorkCompleteDate IS Null "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,1
'3th          j=1
'3th          DtTbl(0)(14) = 0
'3th          For i=1 To Num
'3th            If DtTbl(i)(8)  <> "済" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(j)(13)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "8" '照会
          WriteLogH "b107", "実搬出事前情報一覧", "01",""
          Get_Data Num2,DtTbl
        'エラートラップ開始
          on error resume next
        'DB接続
          dim StrSQL
          ConnDBH ObjConn, ObjRS
          For i=1 To Num2
'CW-002            If DtTbl(i)(13) <> 0 Then
            If DtTbl(i)(13) <> 0 AND DtTbl(i)(6)="" AND DtTbl(i)(14)="未" Then
              StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                       "UpdtTmnl='"& USER &"', TruckerFlag"& DtTbl(i)(13) &"=1 "&_
                       "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                       "WHERE WkNo='"& DtTbl(i)(3) &"' AND WkType='1' AND Process='R' )"
              ObjConn.Execute(StrSQL)
              if err <> 0 then
                Set ObjRS = Nothing
                jampErrerPDB ObjConn,"2","b107","01","実搬出:紹介済処理","104","SQL:<BR>"&strSQL
              end if
            End If
          Next
        'DB接続解除
          DisConnDBH ObjConn, ObjRS
        'エラートラップ解除
          on error goto 0
          Response.Redirect "./dmo010L.asp"
      Case else '全件表示
          WriteLogH "b101", "実搬出事前情報一覧", "04",""
          strWrer = " "
          getData DtTbl,strWrer,0
  End Select

'データ取得関数
'3th chage Function getData(DtTbl,strWhere)
'2009/02/25 Add-S G.Ariola
Function getSort(Key,SortKey,str)
	getSort = str
	
	
'		if str = "" then
'			'getSort = " ORDER BY (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
'			if (FieldName(Key) = "WorkDate" OR FieldName(Key) = "FreeTime" OR FieldName(Key) = "WorkCompleteDate") AND SortKey = "ASC" then 
'			getSort = " ORDER BY isnull(" & FieldName(Key) & ",DATEADD(Year,100,getdate())) " & SortKey	
'			else
'			getSort = " ORDER BY " & FieldName(Key) & " " & SortKey	
'			end if
'			
'		else
'			'getSort = str & " , (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
'			if (FieldName(Key) = "WorkDate" OR FieldName(Key) = "FreeTime" OR FieldName(Key) = "WorkCompleteDate") AND SortKey = "ASC"  then 
'			getSort = str & " , isnull(" & FieldName(Key) & ",DATEADD(Year,100,getdate())) " & SortKey	
'			else
'			getSort = str & " , " & FieldName(Key) & " " & SortKey	
'			end if
'			
'		end if	
	if str = "" AND Key<>"" then
		str = " ORDER BY "
	elseif str <> "" AND Key<>"" Then 
		str = str & ","	
	elseif str = "" AND Key = "" then
		str = " ORDER BY WorkDate_Sort ASC ,InputDate ASC "		
	end if

	if Key <> "" then 
		if FieldName(CInt(Key)) = "FreeTime" AND SortKey = "ASC" then 
			str = str & " FreeTime_Sort ASC "
		elseif FieldName(CInt(Key)) = "WorkCompleteDate" AND SortKey = "ASC" then 
			str = str & " WorkCompleteDate_Sort ASC "
		elseif FieldName(CInt(Key)) = "WorkDate" AND SortKey = "ASC" then 
			str = str & " WorkDate_Sort ASC "
		else
			str = str & FieldName(Key) & " " & SortKey	
		end if			
	end if
	
	getSort = str	
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
Function getData(DtTbl,strWhere,DelType)

  ReDim DtTbl(1)
'C-002  DtTbl(0)=Array("搬出先","搬出<BR>予定日","指示元","作業<BR>番号","指定種類","コンテナ番号<BR>／ＢＬ番号","完了日時","返却予定","返却","指示先","指示先<BR>回答","BL番号","アラームフラグ","照会先","指示元へ回答","船社","船名","サイズ","ＣＹ","フリー<BR>タイム","搬出完了日","返却値")
'3th  DtTbl(0)=Array("搬出先","搬出<BR>予定日","指示元","作業<BR>番号","指定種類","コンテナ番号<BR>／ＢＬ番号","完了日時","返却予定","返却","指示先","指示先<BR>回答","BL番号","アラームフラグ","照会先","指示元へ回答","船社","船名","サイズ","ＣＹ","フリー<BR>タイム","搬出完了日","返却値","備考１","備考２","備考３")
'Chang 20050303 STAT fro 4th Recon By SEIKO N.Oosige
'  DtTbl(0)=Array("搬出先","搬出<BR>予定日","指示元","作業<BR>番号","指定種類","コンテナ番号<BR>／ＢＬ番号","完了日時","返却予定","返却","指示先","指示先<BR>回答","BL番号","アラームフラグ","照会先","指示元へ回答","船社","船名","サイズ","ＣＹ","フリー<BR>タイム","搬出完了日","返却値","備考１","備考２","納入先１")
  DtTbl(0)=Array("搬出先","搬出<BR>予定日","指示元","作業<BR>番号","指定種類","コンテナ番号<BR>／ＢＬ番号","完了日時","返却予定","返却","指示先","指示先<BR>回答","BL番号","アラームフラグ","照会先","指示元へ回答","船社","船名","SZ","ＣＹ","フリー<BR>タイム","搬出完了日","返却値","備考１","備考２","納入先１","コード","指示元<BR>担当")
'Chang 20050303 END
'2009/02/25 Add-S G.Ariola
dim ctr
for ctr = 1 to 3
Session(CSTR("Key" & ctr))
if Session(CSTR("Key" & ctr)) <> "" then
	Select Case Session(CSTR("Key" & ctr))
		Case "0" '搬入予定日
			DtTbl(0)(1) = DtTbl(0)(1) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "1" '指示元 － コード
			DtTbl(0)(25) = DtTbl(0)(25) & getImage(Session(CSTR("KeySort" & ctr)))		
		Case "2" '作業番号
			DtTbl(0)(3) = DtTbl(0)(3) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "3" '指定種類
			DtTbl(0)(4) = DtTbl(0)(4) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "4" 'コンテナ番号/BL番号
			DtTbl(0)(5) = DtTbl(0)(5) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "5" '船社
			DtTbl(0)(15) = DtTbl(0)(15) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "6" '船名
			DtTbl(0)(16) = DtTbl(0)(16) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "7" 'SZ
			DtTbl(0)(17) = DtTbl(0)(17) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "8" 'ＣＹ
			DtTbl(0)(18) = DtTbl(0)(18) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "9" 'フリータイム
			DtTbl(0)(19) = DtTbl(0)(19) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "10" '納入先１
			DtTbl(0)(24) = DtTbl(0)(24) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "11" '完了日時
			DtTbl(0)(6) = DtTbl(0)(6) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "12" '返却予定
			DtTbl(0)(7) = DtTbl(0)(7) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "13" '返却
			DtTbl(0)(8) = DtTbl(0)(8) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "14" '指示先 － コード
			DtTbl(0)(9) = DtTbl(0)(9) & getImage(Session(CSTR("KeySort" & ctr)))
'		Case "16" '指示先 － 担当
'			DtTbl(0)(28) = DtTbl(0)(28) & getImage(Session(CSTR("KeySort" & ctr)))		
		Case "15" '指示先回答
			DtTbl(0)(10) = DtTbl(0)(10) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "16" '備考１
			DtTbl(0)(22) = DtTbl(0)(22) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "17" '備考２
			DtTbl(0)(23) = DtTbl(0)(23) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "18" '指示元 － 担当
			DtTbl(0)(26) = DtTbl(0)(26) & getImage(Session(CSTR("KeySort" & ctr)))
	  End Select
end if	  
next
'2009/02/25 Add-E G.Ariola
'3th Add Start
  Dim DelStr,DelTarget
  DelStr=Array("","済","未","No")
  DelTarget=Array(0,8,10,10)
  DtTbl(0)(14) = 0
'3th Add End
  'エラートラップ開始
    on error resume next
  'DB接続
    dim StrSQL
    ConnDBH ObjConn, ObjRS
  '対象件数取得
    StrSQL = "SELECT count(WkContrlNo) AS CNUM FROM hITCommonInfo ITC "&_
             "WHERE WkType='1' AND (RegisterCode='"& USER &"' "&_
             "OR TruckerSubCode1='"& COMPcd &"' OR TruckerSubCode2='"&_
              COMPcd &"' OR TruckerSubCode3='"& COMPcd &"' OR TruckerSubCode4='"& COMPcd &"') AND Process='R' " &_
              strWhere
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "2","b101","00","実搬出:一覧表示(対象件数取得)","101","SQL:<BR>"&strSQL
      Exit Function
    end if
    Num = ObjRS("CNUM")
    ObjRS.close

  'データ取得 '2009/11/02 Tanaka 初期表示ソート用に,ITC.InputDate を追加
    StrSQL ="SELECT T.* FROM (SELECT ITC.DeliverTo,ITC.BLNo, "&_
			"ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate, "&_
			"CASE WHEN ITC.WorkDate Is NULL THEN '9999/12/31' WHEN ITC.WorkDate ='' THEN '9999/12/31' ELSE ITC.WorkDate END as WorkDate_Sort, "&_
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
			"ITC.DeliverTo1,"&_
			"ISNULL(CONVERT(varchar(10),ITC.WorkCompleteDate,111),'') as WorkCompleteDate, "&_
			"CASE WHEN ITC.WorkCompleteDate Is NULL THEN '9999/12/31' WHEN ITC.WorkCompleteDate ='' THEN '9999/12/31' ELSE ITC.WorkCompleteDate END as WorkCompleteDate_Sort, "&_			
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
			"A.CYDelTime, INC.ReturnTime,ITC.InputDate "&_
			"FROM hITCommonInfo ITC "&_
			"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
			"LEFT JOIN ImportCont AS INC ON (ITC.ContNo=INC.ContNo) "&_
			"LEFT JOIN mVessel AS mV On (INC.VslCode=mV.VslCode) "&_
			"LEFT JOIN (SELECT Distinct BL.BLNo, INC.FreeTime, MIN(INC.CYDelTime) AS CYDelTime,mV.ShipLine, mV.FullName, BL.RecTerminal "&_

			"FROM ImportCont AS INC "&_
			"LEFT JOIN mVessel AS mV ON INC.VslCode = mV.VslCode "&_
			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
			"GROUP BY BL.BLNo,INC.FreeTime,mV.ShipLine, mV.FullName, BL.RecTerminal) A ON  A.BLNo=ITC.BLNo "&_
			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType='4' AND (ITC.RegisterCode='"& USER &"' "&_
			"OR ITC.TruckerSubCode1='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode2='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode3='"& COMPcd &"' "&_
			"OR ITC.TruckerSubCode4='"& COMPcd &"')  "& strWhere &" "&_
			"UNION ALL "&_
			"SELECT ITC.DeliverTo,ITC.BLNo, "&_
			"ISNULL(CONVERT(varchar(10),ITC.WorkDate,111),'') as WorkDate, "&_
			"CASE WHEN ITC.WorkDate Is NULL THEN '9999/12/31' WHEN ITC.WorkDate ='' THEN '9999/12/31' ELSE ITC.WorkDate END as WorkDate_Sort, "&_
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
			"CASE WHEN INC.FreeTime Is NULL THEN '9999/12/31' WHEN INC.FreeTime ='' THEN '9999/12/31' ELSE INC.FreeTime END as FreeTime_Sort, "&_
			"ITC.DeliverTo1, "&_			
			"ISNULL(CONVERT(varchar(10),ITC.WorkCompleteDate,111),'') as WorkCompleteDate, "&_
			"CASE WHEN ITC.WorkCompleteDate Is NULL THEN '9999/12/31' WHEN ITC.WorkCompleteDate ='' THEN '9999/12/31' ELSE ITC.WorkCompleteDate END as WorkCompleteDate_Sort, "&_		
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
			"INC.CYDelTime, INC.ReturnTime,ITC.InputDate "&_
			"FROM hITCommonInfo ITC "&_
			"LEFT JOIN hITReference ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
			"INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
			"LEFT JOIN ImportCont AS INC ON ITC.ContNo=INC.ContNo "&_
			"LEFT JOIN mVessel AS mV On INC.VslCode=mV.VslCode "&_
			"LEFT JOIN BL ON INC.VslCode=BL.VslCode AND INC.VoyCtrl=BL.VoyCtrl AND INC.BLNo=BL.BLNo "&_
			"LEFT JOIN Container AS CON ON INC.VslCode = CON.VslCode AND INC.VoyCtrl=CON.VoyCtrl AND INC.ContNo=CON.ContNo "&_
			"WHERE ITC.Process='R' AND WkType='1' AND ITC.FullOutType<>'4' AND (ITC.RegisterCode='"& USER &"' "&_
			"OR ITC.TruckerSubCode1='"& COMPcd &"' OR ITC.TruckerSubCode2='"& COMPcd &"' OR ITC.TruckerSubCode3='"& COMPcd &"' OR ITC.TruckerSubCode4='"& COMPcd &"') "& strWhere &") AS T "&_
             strOrder
			 '"ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
'C-002 ADD This Line : "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
'C-004 ADD This Line : "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
'3th chage This Line : "ITC.Comment1, ITC.Comment2, ITC.Comment3, "&_
'3th add INC.VoyCtrl, INC.VslCode,
'response.Write(StrSQL)
'response.end
	'2010/09/30 C.Pestano Upd-S
	ObjRS.PageSize = 200
	ObjRS.CacheSize = 200
	ObjRS.CursorLocation = 3

    ObjRS.Open StrSQL, ObjConn
	Num2 = ObjRS.recordcount
	'Y.TAKAKUWA Add-S 2014-11-19 
	'Response.write StrSQL
	'Response.end
	'Y.TAKAKUWA Add-E 2014-11-19
	ReDim Preserve DtTbl(Num2)
	'abspage = 1
	'pagecnt = 1
		
	if CInt(Num2) > ObjRS.PageSize then		
		If CInt(Request("pagenum")) = 0 Then
			ObjRS.AbsolutePage = 1
		Else
			If CInt(Request("pagenum")) <= ObjRS.PageCount Then
				ObjRS.AbsolutePage = CInt(Request("pagenum"))				
			Else
				ObjRS.AbsolutePage = 1	
			End If			
		End If		
		abspage = ObjRS.AbsolutePage
		pagecnt = ObjRS.PageCount
	End If	

	'2010/09/30 C.Pestano Upd-E	
	
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "2","b101","00","実搬出:一覧表示(データ取得)","102","SQL:<BR>"&strSQL
      Exit Function
    end if
    dim tmpBLNo(1),tmptime
    tmpBLNo(0) = ""
    tmpBLNo(1) = ""
    i=1
	RecCtr = 0	
    Do Until ObjRS.EOF	 
	 '2015-01-23 Y.TAKAKUWA Upd-S
	 'if RecCtr <= ObjRS.PageSize then
     if RecCtr <= ObjRS.PageSize - 1 then	 
	 '2015-01-23 Y.TAKAKUWA Upd-E
     If DtTbl(i-1)(3)<>Trim(ObjRS("WkNo")) Then
'C-002      DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","")
      'DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","","","","")
	  DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","","","","","","")
      DtTbl(i)(0)=Trim(ObjRS("DeliverTo"))

      DtTbl(i)(1)=Mid(ObjRS("WorkDate"),3,8)
      DtTbl(i)(3)=Trim(ObjRS("WkNo"))
      DtTbl(i)(4)=Trim(ObjRS("FullOutType"))
      DtTbl(i)(5)=Trim(ObjRS("BLContNo"))
      DtTbl(i)(6)=Trim(Mid(ObjRS("WorkCompleteDate"),3,14))
      If Trim(Mid(DtTbl(i)(6),10))<>"" Then
        tmptime=Split(Mid(DtTbl(i)(6),10),":",3,1)
        DtTbl(i)(6)=Left(DtTbl(i)(6),9)&Right(0&tmptime(0),2)&":"&tmptime(1)
      End If
      DtTbl(i)(7)=Trim(ObjRS("ReturnDateStr"))
'chenge 030722      DtTbl(i)(12)=ObjRS("ReturnArrert")
      DtTbl(i)(8)=ObjRS("ReturnValue")
      DtTbl(i)(21)=Trim(ObjRS("ReturnDateVal"))
'      If DtTbl(i)(4) = 4 Then
'        DtTbl(i)(5)=Trim(ObjRS("BLNo"))
'        DtTbl(i)(11)=Trim(ObjRS("BLNo"))
'      ElseIf DtTbl(i)(4) = 2 Then
'        DtTbl(i)(5)=Trim(ObjRS("ConTNo"))
        DtTbl(i)(11)=Trim(ObjRS("BLNo"))
'      Else
'        DtTbl(i)(5)=Trim(ObjRS("ConTNo"))
'      End If
	  
'2009/02/25 Add-S G.Ariola		
		DtTbl(i)(2) = Trim(ObjRS("Code1"))
		DtTbl(i)(25) = Trim(ObjRS("Name1"))
		DtTbl(i)(9) = Trim(ObjRS("Code2"))
		DtTbl(i)(26) = Trim(ObjRS("Name2"))

		DtTbl(i)(10) = ObjRS("Flag1")
		DtTbl(i)(14) = ObjRS("Flag2")

If Trim(ObjRS("TruckerSubCode4")) = COMPcd Then
DtTbl(i)(13) = 4
ElseIf Trim(ObjRS("TruckerSubCode3")) = COMPcd Then
DtTbl(i)(13) = 3
ElseIf Trim(ObjRS("TruckerSubCode2")) = COMPcd Then
DtTbl(i)(13) = 2
ElseIf Trim(ObjRS("TruckerSubCode1")) = COMPcd Then
DtTbl(i)(13) = 1
Else
DtTbl(i)(13) = 0
end if
'2009/02/25 Add-E G.Ariola	  
'2009/02/25 Del-S G.Ariola	  
'   '指示先照会済みフラグ
'      If Trim(ObjRS("TruckerSubCode4")) = COMPcd Then
'        DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode3"))
'        DtTbl(i)(9) = Null
'        DtTbl(i)(13) = 4
'        DtTbl(i)(14) = ObjRS("TruckerFlag4")
'      ElseIf Trim(ObjRS("TruckerSubCode3")) = COMPcd Then
'        DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode2"))
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode4"))
'        DtTbl(i)(13) = 3
'        DtTbl(i)(14) = ObjRS("TruckerFlag3")
'        DtTbl(i)(10) = ObjRS("TruckerFlag4")
'      ElseIf Trim(ObjRS("TruckerSubCode2")) = COMPcd Then
'        DtTbl(i)(2) = Trim(ObjRS("TruckerSubCode1"))
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode3"))
'        DtTbl(i)(13) = 2
'        DtTbl(i)(14) = ObjRS("TruckerFlag2")
'        DtTbl(i)(10) = ObjRS("TruckerFlag3")
'      ElseIf Trim(ObjRS("TruckerSubCode1")) = COMPcd Then
'        If ObjRS("UserType") = "5" Then         'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("HeadCompanyCode"))  'CW-051
'        Else                        'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("RegisterCode"))
'        End If                      'CW-051
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode2"))
'        DtTbl(i)(13) = 1
'        DtTbl(i)(14) = ObjRS("TruckerFlag1")
'        DtTbl(i)(10) = ObjRS("TruckerFlag2")
'      Else
'        If ObjRS("UserType") = "5" Then         'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("HeadCompanyCode"))  'CW-051
'        Else                        'CW-051
'          DtTbl(i)(2) = Trim(ObjRS("RegisterCode"))
'        End If                      'CW-051
'        DtTbl(i)(9) = Trim(ObjRS("TruckerSubCode1"))
'        DtTbl(i)(13) = 0
'        DtTbl(i)(14) = Null
'        DtTbl(i)(10) = ObjRS("TruckerFlag1")
'      End If
'2009/02/25 Del-E G.Ariola	  
	  
'      If IsNull(DtTbl(i)(9)) Then
'        DtTbl(i)(10) ="　"
'      ElseIf DtTbl(i)(10) = 0 Then
'        DtTbl(i)(10) ="未"
'      ElseIf DtTbl(i)(10) = 1 Then
'        DtTbl(i)(10) ="Yes"
'      Else
'        DtTbl(i)(10) ="No"
'      End If
'      If DtTbl(i)(14)=0 Then
'        DtTbl(i)(14) ="未"
'      ElseIf DtTbl(i)(14) = 1 Then
'        DtTbl(i)(14) ="Yes"
'      ElseIf DtTbl(i)(14) = 2 Then
'        DtTbl(i)(14) ="No"
'      Else
'        DtTbl(i)(14) ="　"
'      End If

'2009/02/25 Del-S G.Ariola	  
'      If DtTbl(i)(4) = 1  Then
'        If IsNull(ObjRS("ReturnTime")) Then
'          DtTbl(i)(8)="未"
'        Else
'          DtTbl(i)(8)="済"
'        End If
'      Else
'        DtTbl(i)(8)="　"
'      End If
'2009/02/25 Del-E G.Ariola	  	  
	  
      DtTbl(i)(12)="-"
'      If DtTbl(i)(4) = 4  Then
'        tmpBLNo(0) = tmpBLNo(0) &","& i  
'        tmpBLNo(1) = tmpBLNo(1) &",'"& DtTbl(i)(5) & "'"
'        DtTbl(i)(17)=""
'        DtTbl(i)(15)=""
'        DtTbl(i)(16)=""
'        DtTbl(i)(18)=""
'        DtTbl(i)(19)=""
'        DtTbl(i)(20)=""
'      Else
        DtTbl(i)(17)=Trim(ObjRS("ContSize"))
        DtTbl(i)(15)=Trim(ObjRS("ShipLine"))
'C-001        DtTbl(i)(16)=Left(ObjRS("FullName"),12)
        DtTbl(i)(16)=Trim(ObjRS("ShipName"))
        DtTbl(i)(18)=Trim(ObjRS("CY"))
        DtTbl(i)(19)=Mid(ObjRS("FreeTime"),3,8)
        DtTbl(i)(20)=Trim(ObjRS("CYDelTime"))
'      End If
'C-002 ADD START
      DtTbl(i)(22)= Trim(ObjRS("Comment1"))
      DtTbl(i)(23)= Trim(ObjRS("Comment2"))
'3th Chage Start
'3th'      DtTbl(i)(24)= Trim(ObjRS("Comment3"))
      DtTbl(i)(24)= Trim(ObjRS("DeliverTo1"))
'3th Change End
'C-002 ADD END

'3th Add Start
      If DelType=0 OR (DelType=1 AND DtTbl(i)(DelTarget(DelType)) <> DelStr(DelType)) OR ((DelType=2 OR DelType=3) AND DtTbl(i)(DelTarget(DelType)) = DelStr(DelType)) Then
        DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(i)(13)
        i=i+1
      Else
        Num=Num-1
      End If
'      If (DelType=1 AND DtTbl(i)(DelTarget(DelType)) = DelStr(DelType)) OR ((DelType=2 OR DelType=3) AND DtTbl(i)(DelTarget(DelType)) <> DelStr(DelType)) Then
'        Num=Num-1
'      Else
'        DtTbl(0)(14) = DtTbl(0)(14) + DtTbl(i)(13)
'        i=i+1
'      End If
'      i=i+1
'3th Add End
     End If 
	 RecCtr = RecCtr + 1 
	 end if
     ObjRS.MoveNext
    Loop   
	
    If i - 1 < Num Then
      'ErrerM = "<DIV class=alert>登録データのうち"& Num2-i+1 &"件について関連データ取得失敗のため"&_
      '         "表示されていません。<BR>システム管理者に問い合わせてください。</DIV><P>"
      Num2 = i - 1
'        DisConnDBH ObjConn, ObjRS
'        jampErrerP "2","b101","00","実搬出:一覧表示","106",Num-i+1&"件のデー"
    ElseIf i > Num Then     'CW-325 ADD
      Num = i - 1           'CW-325 ADD
    End If
    '2010/04/23 M.Marquez Add-S
    if Ubound(DtTbl) < Num Then 
        Num=Ubound(DtTbl)
    end if
    '2010/04/23 M.Marquez Add-E
'    If Err <> 0 Then
'      DisConnDBH ObjConn, ObjRS 'DB切断
'      jampErrerP "2","b101","00","実搬出:一覧表示(データ編集)","200",i&"番目のデータ編集エラー"
'      Exit Function
'    End If
'
''ADD 20030729 BL指定の場合の追加データ取得 Start
'    If tmpBLNo(1) <> "" Then
'      StrSQL = "SELECT INC.BLNo, INC.FreeTime,INC.CYDelTime, mV.ShipLine, mV.FullName, BL.RecTerminal "&_
'               "FROM ImportCont AS INC LEFT JOIN mVessel AS mV ON INC.VslCode = mV.VslCode "&_
'               "LEFT JOIN BL ON (INC.BLNo=BL.BLNo) AND (INC.VoyCtrl=BL.VoyCtrl) AND (INC.VslCode=BL.VslCode) "&_
'               "WHERE INC.BLNo IN("& Mid(tmpBLNo(1),2) &") ORDER BY INC.BLNo,INC.UpdtTime DESC"
''3th add INC.VoyCtrl, INC.VslCode,
'      ObjRS.Open strSQL, ObjConn
'      If Err <> 0 Then
'        DisConnDBH ObjConn, ObjRS
'        jampErrerP "2", "b101", "00", "実搬出:一覧表示(追加項目取得)", "101", "SQL:<BR>" & strSQL
'      End If
'      Dim tmpBLNoA(1)
'      tmpBLNoA(0) = Split(tmpBLNo(0), ",", -1, 1)
'      tmpBLNoA(1) = Split(tmpBLNo(1), ",", -1, 1)
'      tmpBLNo(1) = ""
'      Do Until ObjRS.EOF
'        If tmpBLNo(1) <> Trim(ObjRS("BLNo")) Then
'          For i = 1 To UBound(tmpBLNoA(0))
'            If tmpBLNoA(1)(i) = "'" & Trim(ObjRS("BLNo")) & "'" Then
'              DtTbl(tmpBLNoA(0)(i))(15) = Trim(ObjRS("ShipLine"))
''C-001              DtTbl(tmpBLNoA(0)(i))(16)=Left(ObjRS("FullName"),12)
'              DtTbl(tmpBLNoA(0)(i))(16) = Trim(ObjRS("FullName"))
'              DtTbl(tmpBLNoA(0)(i))(18) = Trim(ObjRS("RecTerminal"))
'              DtTbl(tmpBLNoA(0)(i))(19) = Mid(ObjRS("FreeTime"), 3, 8)
'              DtTbl(tmpBLNoA(0)(i))(20) = Trim(ObjRS("CYDelTime"))
'              tmpBLNo(1) = Trim(ObjRS("BLNo"))
'            End If
'         Next
'       End If
'       ObjRS.MoveNext
'      Loop
'      ObjRS.Close
'      If Err <> 0 Then
'        DisConnDBH ObjConn, ObjRS   'DB切断
'        jampErrerP "2","b101","00","実搬出:一覧表示(追加項目データ編集)","200",i&"番目のデータ編集エラー"
'        Exit Function
'      End If
'   End If

'ADD 20030729 BL指定の場合の追加データ取得 End
  'DB接続解除
    'DisConnDBH ObjConn, ObjRS
  'エラートラップ解除
    on error goto 0
End Function


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>実搬出情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//データが無い場合の表示制御
function vew(){
<%If Num2<>0 Then%>
    //Y.TAKAKUWA Add-S 2015-03-06
	var IEVersion = getInternetExplorerVersion();
	//Y.TAKAKUWA Add-E 2015-03-06
	
	var obj3=document.getElementById("BDIV");
	
	// Y.TAKAKUWA Upd-S 2015-03-09
	//if((document.body.offsetWidth-200)  < 880){
	//	obj3.style.width=document.body.offsetWidth;
	//	obj3.style.overflowX="auto";
	//}
	//else{
	 <% If DtTbl(0)(14)<>0 Then %>
    //  obj3.style.width=document.body.offsetWidth-120;
	 <% Else %>
	//	obj3.style.width=document.body.offsetWidth-79;
	 <% End If %>
	//	obj3.style.overflowX="auto";
	//}
	
	if(IEVersion < 10)
	{
		if((document.body.offsetWidth-200) < 880) {
			obj3.style.width=document.body.offsetWidth;
			obj3.style.overflowX="auto";
		}
		else{
			<% If DtTbl(0)(14)<>0 Then %>
			obj3.style.width=document.body.offsetWidth-120;
			<% Else %>
			obj3.style.width=document.body.offsetWidth-79;
			<% End If %>
			obj3.style.overflowX="auto";
		}
		//Y.TAKAKUWA Upd-S 2015-03-09
		obj3.style.height=document.body.offsetHeight-20;
		obj3.style.overflowY="scroll";
		//Y.TAKAKUWA Upd-E 2015-03-09
	}
	else{
	    var initialHeight = document.documentElement.clientHeight;
		if((document.body.offsetWidth-200) < 880) {
			obj3.style.width=document.body.offsetWidth-25;
			obj3.style.overflowX="auto";
		}
		else{
			<% If DtTbl(0)(14)<>0 Then %>
			obj3.style.width=document.body.offsetWidth-205;
			<% Else %>
			obj3.style.width=document.body.offsetWidth-79;
			<% End If %>
			obj3.style.overflowX="auto";
		}	
		//Y.TAKAKUWA Upd-S 2015-03-09
		obj3.style.height=initialHeight-70;
		obj3.style.overflowY="scroll";
		//Y.TAKAKUWA Upd-E 2015-03-09
	}
	// Y.TAKAKUWA Upd-E 2015-03-09
	
	//Y.TAKAKUWA Add-S 2015-03-06
	var obj3header=document.getElementById("BDIVHEADER");
	if(IEVersion < 10)
	{
		if((document.body.offsetWidth-200) < 880) {
			obj3header.style.width=document.body.offsetWidth;
		}
		else{
			<% If DtTbl(0)(14)<>0 Then %>
			obj3header.style.width=document.body.offsetWidth-120;
			<% Else %>
			obj3header.style.width=document.body.offsetWidth-79;
			<% End If %>
		}	
		obj3header.style.height = 35;
	}
	else
	{
		if((document.body.offsetWidth-200) < 880) {
			obj3header.style.width=document.body.offsetWidth-40;
		}
		else{
			<% If DtTbl(0)(14)<>0 Then %>
			obj3header.style.width=document.body.offsetWidth-220;
			<% Else %>
			obj3header.style.width=document.body.offsetWidth-79;
			<% End If %>
		}
		obj3header.style.height = 35;
	}
	//Y.TAKAKUWA Add-S 2015-03-06
	
<% End If %>
}
//Y.TAKAKUWA Add-S 2015-03-06
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
//Y.TAKAKUWA Add-E 2015-03-06
//更新
function GoRenew(No){
  Fname=document.dmo010F;
  Fname.targetNo.value=No;
  Fname.action="./dmi015.asp";
  newWin = window.open("", "ReEntry", "left=10,top=10,status=yes,resizable=yes,scrollbars=yes");
  Fname.target="ReEntry";
  Fname.elements[No].disabled=false;
  Fname.submit();
  Fname.elements[No].disabled=true;
  Fname.target="_self";
}
//コンテナ詳細
function GoConinf(wkcNo,flag,conNo){
  Fname=document.dmo010F;
  Fname.SakuNo.value=wkcNo;
  Fname.flag.value=flag;
  Fname.CONnum.value=conNo;
//  Fname.InfoFlag.value="9";
//  Fname.action="./dmi015.asp";
  Fname.action="./dmo900.asp";
  newWin = window.open("", "ConInfo", "left=30,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes");
  Fname.target="ConInfo";
  Fname.submit();
  Fname.target="_self";
  Fname.InfoFlag.value="0";
}
//検索
function SerchC(SortFlag,Kye){
  Fname=document.dmo010F;
  Fname.SortFlag.value=SortFlag;
  Fname.SortKye.value=Kye;
  Fname.target="_self";
  Fname.action="./dmo010L.asp";
  Fname.submit();
}
//照会済
function GoSyokaizumi(){
  target=document.dmo010F;
  if(target.DataNum.value>0){
    flag = confirm('未回答の回答を「Yes」にしますか？');
    if(flag==true){
      len=target.elements.length;
      for(i=0;i<len;i++){
        target.elements[i].disabled=false;
      }
      target.SortFlag.value=8;
      target.target="_self";
      target.action="./dmo010L.asp";
      target.submit();
    }
  }
}
//展開
function GoDevelop(No){
  Fname=document.dmo010F;
  Fname.targetNo.value=No;
  Fname.action="./dmo030.asp";
  newWin = window.open("", "ConInfo", "left=50,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes");
  Fname.target="ConInfo";
  Fname.elements[0].disabled=false;
  Fname.elements[No].disabled=false;
  Fname.submit();
  Fname.elements[0].disabled=true;
  Fname.elements[No].disabled=true;
  Fname.target="_self";
}
//CSV		ADD C-001
function GoCSV(){
  target=document.dmo010F;
  target.target="Bottom";
  len=target.elements.length;
  for(i=0;i<len;i++){
    target.elements[i].disabled=false;
  }
  target.action="./dmo080.asp";
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
//Y.TAKAKUWA Add-S 2015-01-22
function showPage(pageNo)
{
   var url = window.location.pathname;
   var filename = url.substring(url.lastIndexOf('/')+1);
   target=document.dmo010F;
   len=target.elements.length;
   for(i=0;i<len;i++){
     target.elements[i].disabled=false;
   }
   target.target="_self";
   filename = "./" + filename
   target.action=filename;
   document.forms[0].pagenum.value=pageNo;
   document.forms[0].submit();
   return false;
}
//Y.TAKAKUWA Add-E 2015-01-22
//Y.TAKAKUWA Add-S 2015-03-06
function cloneTable(tblSource, tblDestination)
{
    <%If Num2<>0 Then%>
	var source = document.getElementById(tblSource);
	var destination = document.getElementById(tblDestination);
	var copy = source.cloneNode(true);
	copy.setAttribute('id', '');
	//Y.TAKAKUWA Add-S 2015-04-06
	//Change the name of cloned elements
	var rowCount = copy.rows.length;
	for(var i=0; i<rowCount; i++) {
		var row = copy.rows[i];
		element_i = row.getElementsByTagName ('input')[0];
		//Y.TAKAKUWA Upd-S 2015-04-08
		// Remove the element from the cloned table.
		//element_i.removeAttribute('name');
		element_i.parentNode.removeChild(element_i);
		//Y.TAKAKUWA Upd-E 2015-04-08
	}
	//Y.TAKAKUWA Add-E 2015-04-06
	destination.parentNode.replaceChild(copy, destination);
	source.style.marginTop = "-35px";
	<%End If%>
}
function onScrollDiv(Scrollablediv,Scrolleddiv) {
    document.getElementById(Scrolleddiv).scrollLeft = Scrollablediv.scrollLeft;
}
//Y.TAKAKUWA Add-E 2015-03-06

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
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" onLoad="vew();" onResize="vew()">
<!--setTimeout('showContent()', 500); -->
<!-------------実搬出情報一覧画面List--------------------------->
<!--<div id="content" style="display:none;"> -->
<%=ErrerM%>
<Form name="dmo010F" method="POST">
<TABLE border="0" cellPadding="2" cellSpacing="0" width="100%">
  <tr>
    <!--Y.TAKAKUWA Upd-S 2015-03-11-->
    <!--<td align="right">-->
	<td align="right" style="min-width:200px;">
	<!--Y.TAKAKUWA Upd-E 2015-03-11-->
	<% 	if Num > 0 then
	        '2015-01-23 Y.TAKAKUWA Upd-S
			'call gfPutPageSort(Num,abspage,pagecnt,"pagenum",SortFlag)
			call gfPutPageSort2(Num,abspage,pagecnt,"pagenum",SortFlag)
			'2015-01-23 Y.TAKAKUWA Upd-E			
		end if
		ObjRS.close	
	%>
	</td>
	<td width="50"></td>	
  </tr>
</TABLE>
<!--Y.TAKAKUWA Add-S 2015-03-05-->	
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
	<td>
		<DIV ID="BDIVHEADER" style="overflow:hidden;">
			<table border="1" cellpadding="0" cellspacing="0" width="100%" Id="testt1">											
			</table>
		</DIV>
	</td>
</tr>
<tr>
<td>
<!--Y.TAKAKUWA Add-E 2015-03-05-->	
<div id="BDIV" onscroll="onScrollDiv(this,'BDIVHEADER');">
<TABLE id="testt" border="1" cellPadding="2" cellSpacing="0">
<%If Num>0 Then%> 
  <% If DtTbl(0)(14)<>0 Then %>
  <THEAD>
  <TR>
    <TH class="hlist" nowrap><%=DtTbl(0)(1)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(2)%><BR><%=DtTbl(0)(25)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(24)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(6)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(7)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(9)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(10)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(22)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(23)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(26)%>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>,<%=DtTbl(0)(20)%>,<%=DtTbl(0)(21)%>,<%=DtTbl(0)(22)%>,<%=DtTbl(0)(23)%>,<%=DtTbl(0)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    </TH>
  </TR>
  </THEAD>
  <TBODY>
    <% For j=1 to i-1 %>
  	<TR class=bgw>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=j%>');"><%=DtTbl(j)(3)%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoDevelop('<%=j%>');"><%=Siji(DtTbl(j)(4))%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(3)%>',<%=DtTbl(j)(4)%>,'<%=DtTbl(j)(5)%>')"><%=DtTbl(j)(5)%></A><BR></TD>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(16),12)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(18),2)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(24),10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(6)%><BR></TD><TD nowrap><%=DtTbl(j)(7)%><BR></TD><TD nowrap><%=DtTbl(j)(8)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(22),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(23),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(25),8)%><BR>
    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>,<%=DtTbl(j)(18)%>,<%=DtTbl(j)(19)%>,<%=DtTbl(j)(20)%>,<%=DtTbl(j)(21)%>,<%=DtTbl(j)(22)%>,<%=DtTbl(j)(23)%>,<%=DtTbl(j)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    </TD>
	</TR>
    <% Next %>
  </TBODY>
  <% Else %>
  <THEAD>  
    <TR>
  <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>,<%=DtTbl(0)(20)%>,<%=DtTbl(0)(21)%>,<%=DtTbl(0)(22)%>,<%=DtTbl(0)(23)%>,<%=DtTbl(0)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    <TH class="hlist" nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%><BR><%=DtTbl(0)(25)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(24)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(6)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(7)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(9)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(10)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(22)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(23)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(26)%></TH>
    </TR>
  </THEAD>  
  <TBODY>
    <% For j=1 to i-1 %>
    <TR class=bgw><INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>,<%=DtTbl(j)(18)%>,<%=DtTbl(j)(19)%>,<%=DtTbl(j)(20)%>,<%=DtTbl(j)(21)%>,<%=DtTbl(j)(22)%>,<%=DtTbl(j)(23)%>,<%=DtTbl(j)(24)%>,<%=Left(DtTbl(j)(25),8)%>'>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=j%>');"><%=DtTbl(j)(3)%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoDevelop('<%=j%>');"><%=Siji(DtTbl(j)(4))%></A><BR></TD>
    <TD nowrap><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(3)%>',<%=DtTbl(j)(4)%>,'<%=DtTbl(j)(5)%>')"><%=DtTbl(j)(5)%></A><BR></TD>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(16),12)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(18),2)%><BR></TD><TD nowrap><%=DtTbl(j)(19)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(24),10)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(6)%><BR></TD><TD nowrap><%=DtTbl(j)(7)%><BR></TD><TD nowrap><%=DtTbl(j)(8)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=DtTbl(j)(10)%><BR></TD>
    <TD nowrap><%=Left(DtTbl(j)(22),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(23),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(25),8)%><BR></TD>
    </TR>
    <% Next %>
  </TBODY>  
  <% End If %>
<% Else %>
  <TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
<% End If 
	DisConnDBH ObjConn, ObjRS
%>
</TABLE>
<!--Y.TAKAKUWA Add-S 2015-03-06-->
<SCRIPT Language="JavaScript">
    cloneTable("testt", "testt1")
</SCRIPT>
 <!--Y.TAKAKUWA Add-E 2015-03-06-->	
</div>
</td>
</tr>
</table>
<%'3th del Set_Data Num,DtTbl %>
<%'2011/05/25 Upd-S Tanaka Num2だとSQLでのデータ取得件数なので指定種類が一覧のデータが存在すると件数が合わなくなりCSV出力でエラーとなるため画面表示件数に変更 %>
<!--  <INPUT type=hidden name=DataNum value="<%=Num2%>"> -->
  <INPUT type=hidden name=DataNum value="<%=i-1%>">
<%'2011/05/25 Upd-E Tanaka  %>
  <INPUT type=hidden name=SortFlag value="<%=SortFlag%>" >
  <INPUT type=hidden name=SortKye value="<%=SortKye%>" >
  <INPUT type=hidden name=InfoFlag value="" >
  <INPUT type=hidden name=SakuNo value="" >
  <INPUT type=hidden name=flag value="" >
  <INPUT type=hidden name=targetNo value="" >
  <INPUT type=hidden name=CONnum value="" >
  <INPUT type=hidden name=pagenum value="" >
  <INPUT type=hidden name=strWhere value="<%=strWrer%>" disabled>
</Form>
<!--</div> -->
<!-------------画面終わり--------------------------->
</BODY></HTML>
