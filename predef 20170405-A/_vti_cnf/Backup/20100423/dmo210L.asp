<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo210L.asp				_/
'_/	Function	:空搬出情報一覧画面リスト出力		_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-001 2003/08/06	CSV出力対応	_/
'_/			:C-002 2003/08/06	備考欄対応	_/
'_/			:B-001 2009/07/14	空コン搬出先不正対応_/
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
'  Session.Contents.Remove("DateP")
'  Session.Contents.Remove("NumP")

'ユーザデータ所得
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")

'INIファイルより設定値を取得
  dim param(2),calcDate2
  getIni param
  calcDate2 = DateAdd("d", "-"&param(2), Date)

'データ取得
  dim Num, DtTbl,i,j,SortFlag,SortKye,ResA

  If Request("SortFlag") = "" Then
    SortFlag = 0
  Else
    SortFlag = Request("SortFlag")
  End If
  ResA=Array("　","Yes","No","未")
  'ソートケース
  dim strWrer
  dim strWrer2		'2009/07/14 Add B-001 Fujiyama 
  
  '2009/02/25 Add-S G.Ariola   
  dim strOrder
  dim FieldName
  ReDim FieldName(18)
  ',"mV.ShipLine"," mV.FullName"
  'FieldName=Array("SPB.InPutDate","Code1","Name1","SPB.BookNo","NumCount","SPB.ContSize1","SPB.ContType1","SPB.ContHeight1","SPB.ContMaterial1","SPB.shipline","mV.FullName","VSLS.CYCut","PickPlace","TruckerCode","SPB.TruckerFlag","SPB.Comment1","SPB.Comment2")
  FieldName=Array("InPutDate","Code1","BookNo","NumCount","ContSize1","ContType1","ContHeight1","ContMaterial1","shipline","FullName","CYCut","PickPlace","TruckerCode","TruckerFlag","Comment1","Comment2","Name1")
   
  strOrder = getSort(Session("Key1"),Session("KeySort1"),"")
  strOrder = getSort(Session("Key2"),Session("KeySort2"),strOrder)
  strOrder = getSort(Session("Key3"),Session("KeySort3"),strOrder)
'2009/02/25 Add-E G.Ariola

  Select Case SortFlag
'2009/02/25 Del-S G.Ariola
      Case "0" '初期表示:入力日順に表示
        WriteLogH "b301", "空搬出事前情報一覧", "01", ""
        strWrer="AND DateDiff(day,SPB.InputDate,'"&calcDate2&"')<=0 "
'3th        getData DtTbl,strWrer
        'GetData DtTbl, strWrer, 0
		GetData DtTbl, strWrer," ", 0
      Case "1" '指示先が未照会のコンテナ一覧
        WriteLogH "b301", "空搬出事前情報一覧", "03", ""
        strWrer="AND DateDiff(day,SPB.InputDate,'"&calcDate2&"')<=0 "
'3th        getData DtTbl,strWrer
        'GetData DtTbl, strWrer, 1
		GetData DtTbl, strWrer," ", 1
'3th        j=1
'3th        DtTbl(0)(6)=0
'3th        For i=1 To Num
'3th         If DtTbl(i)(5) = "未" Then
'3th            DtTbl(j)=DtTbl(i)
'3th            DtTbl(0)(6)=DtTbl(0)(6)+DtTbl(j)(6)
'3th            j=j+1
'3th          End If
'3th        Next
'3th        Num=j-1
      Case "7" '指示先が未照会のコンテナ一覧
        WriteLogH "b301", "空搬出事前情報一覧", "07", ""
        strWrer="AND DateDiff(day,SPB.InputDate,'"&calcDate2&"')<=0 "
'3th        getData DtTbl,strWrer
        'GetData DtTbl, strWrer, 2
		GetData DtTbl, strWrer," ", 2
'3th        j=1
'3th        DtTbl(0)(6)=0
'3th        For i=1 To Num
'3th         If DtTbl(i)(5) = "No" Then
'3th            DtTbl(j)=DtTbl(i)
'3th            DtTbl(0)(6)=DtTbl(0)(6)+DtTbl(j)(6)
'3th            j=j+1
'3th          End If
'3th        Next
'3th        Num=j-1
      Case "2" '搬入未完了分をすべて表示
        WriteLogH "b201", "空搬入事前情報一覧", "02", ""
        strWrer = " "
'3th        getData DtTbl,strWrer
        'GetData DtTbl, strWrer, 0
		GetData DtTbl, strWrer," ", 0
        j = 1
        DtTbl(0)(6) = 0
        For i = 1 To Num
         If DtTbl(i)(7) = "0" Then
            DtTbl(j) = DtTbl(i)
            DtTbl(0)(6) = DtTbl(0)(6) + DtTbl(j)(6)
            j = j + 1
          End If
        Next
        Num = j - 1
      Case "3" '全件表示
        WriteLogH "b301", "空搬出事前情報一覧", "04", ""
        strWrer = " "
'3th        getData DtTbl,strWrer
        'GetData DtTbl, strWrer, 0
		GetData DtTbl, strWrer," ", 0
'2009/02/25 Del-E G.Ariola
      Case "4" 'ブッキング番号で検索
          SortKye=Request("SortKye")
          WriteLogH "b301", "空搬出事前情報一覧","11",SortKye
'          If Session.Contents("ConNum") = "" Then
'            jampErrerP "0","b301","11","空搬出：一覧検索","001",""
'          Else
'            DtTbl=Session("DateT")
'            Num  =Session.Contents("ConNum")
'          End If
'3th chage          Get_Data Num,DtTbl
		  'strWrer = "AND SPB.BookNo LIKE '%" & SortKye & "'"
		  if SortKye <> "" then
          	strWrer = "AND SPB.BookNo LIKE '%" & SortKye & "' AND DateDiff(day,SPB.InputDate,'"&calcDate2&"')<=0 "
          	strWrer2 = "AND A.BookNo LIKE '%" & SortKye & "' AND DateDiff(day,A.InputDate,'"&calcDate2&"')<=0 " 	'2009/07/14 Add B-001 Fujiyama
		  else
		  	strWrer="AND DateDiff(day,SPB.InputDate,'"&calcDate2&"')<=0 "
		  	strWrer2="AND DateDiff(day,A.InputDate,'"&calcDate2&"')<=0 "	'2009/07/14 Add B-001 Fujiyama
		  end if
'2009/07/14 Upd-S B-001 Fujiyama
'          getData DtTbl,strWrer,0
          getData DtTbl,strWrer,strWrer2,0
'2009/07/14 Upd-E B-001 Fujiyama
'3th          j=1
'3th          DtTbl(0)(6)=0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(2),Len(SortKye))= SortKye Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(6)=DtTbl(0)(6)+DtTbl(j)(6)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "8" '照会済
          WriteLogH "b307", "空搬出事前情報照会","01",""
'          If Session.Contents("ConNum") = "" Then
'            jampErrerP "0","b307","01","空搬出：一覧照会","001",""
'          Else
'            DtTbl=Session("DateT")
'            Num  =Session.Contents("ConNum")
'          End If
          Get_Data Num,DtTbl
        'エラートラップ開始
          on error resume next
        'DB接続
          dim ObjConn, ObjRS, StrSQL
          ConnDBH ObjConn, ObjRS
          For i=1 To Num
'CW-002            If DtTbl(i)(5) = "　" Then
            If DtTbl(i)(5) = "　" AND DtTbl(i)(6)=3 AND DtTbl(i)(7)=0 Then
              StrSQL = "UPDATE BookingAssign SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
                       "UpdtTmnl='"& USER &"', TruckerFlag='1' "&_
                       "WHERE BookNo='"& DtTbl(i)(2) &"' AND SenderCode='"& DtTbl(i)(1) &"' "&_
                       "AND TruckerCode='"& DtTbl(i)(12) &"' AND Process='R' "
'3th Change SPBookInfo -> BookingAssign
'3th Dell Status='0', 
'3th TruckerCode='"& DtTbl(i)(12) &"'
              ObjConn.Execute(StrSQL)
              if err <> 0 then
                Set ObjRS = Nothing
                jampErrerPDB ObjConn,"2","b307","01","空搬出：一覧照会","104","SQL:<BR>"&strSQL
              end if
              DtTbl(i)(6)=1
            End If
          Next
        'DB接続解除
          DisConnDBH ObjConn, ObjRS
        'エラートラップ解除
          on error goto 0
'2009/02/25 Add-S G.Ariola  		  
	 Case else '全件表示
          WriteLogH "b101", "実搬出事前情報一覧", "04",""
          strWrer="AND DateDiff(day,SPB.InputDate,'"&calcDate2&"')<=0 "
          strWrer2="AND DateDiff(day,A.InputDate,'"&calcDate2&"')<=0 "	'2009/07/14 Add B-001 Fujiyama
'2009/07/14 Upd-S B-001 Fujiyama
'          getData DtTbl,strWrer,0
          getData DtTbl,strWrer,strWrer2,0
'2009/07/14 Upd-E B-001 Fujiyama
'2009/02/25 Add-E G.Ariola  		  
  End Select
'  Session.Contents.Remove("DateT")
'  Session("DateT")=DtTbl
'  Session.Contents("ConNum")=Num
'  If Num=0 Then
'    Session.Contents("NullFlag")=0
'  Else
'    Session.Contents("NullFlag")=1
'  End If

'2009/02/25 Add-S G.Ariola
Function getSort(Key,SortKey,str)
getSort = str
	if Key <> "" then
	
		if str = "" then
			getSort = " ORDER BY (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
			'getSort = " ORDER BY " & FieldName(Key) & " " & SortKey
		else
			getSort = str & " , (Case When LTRIM(ISNULL(" & FieldName(Key) & "),'')='' Then 1 Else 0 End), " & FieldName(Key) & " " & SortKey
			'getSort = str & " , " & FieldName(Key) & " " & SortKey
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


'データ取得関数
'3th Function getData(DtTbl,strWrer)
'2009/07/14 Upd-S B-001 Fujiyama
'Function getData(DtTbl,strWrer,DelType)
Function getData(DtTbl,strWrer,strWrer2,DelType)
'2009/07/14 Upd-E B-001 Fujiyama
  ReDim DtTbl(1)
'CW-002  DtTbl(0)=Array("入力日","指示元","ブッキング番号","ピック済本数","指示先","指示先照会")
'CW-003  DtTbl(0)=Array("入力日","指示元","ブッキング番号","ピック済本数","指示先","指示先照会","照会フラグ")
'C-002  DtTbl(0)=Array("入力日","指示元","ブッキング番号","ピック済本数","指示先","指示先回答","照会フラグ","作業完了F","船社","船名","指示元表示用")
'3th DtTbl(0)=Array("入力日","指示元","ブッキング番号","ピック済本数","指示先","指示先回答","照会フラグ","作業完了F","船社","船名","指示元表示用","備考１")
'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
'  DtTbl(0)=Array("入力日","指示元","ブッキング番号","ピック済本数","指示先","指示先回答","照会フラグ","作業完了F","船社","船名","指示元表示用","備考１","指示先データ","備考２","サイズ","タイプ","高さ","材質")
  'DtTbl(0)=Array("入力日","指示元","ブッキング番号","ピック済","指示先","指示先<BR>回答","照会フラグ","作業完了F","船社","船名","指示元表示用","備考１","指示先データ","備考２","SZ","タイプ","高さ","材質")
  DtTbl(0)=Array("入力日","指示元","ブッキング番号","ピック済","指示先","指示先<BR>回答","照会フラグ","作業完了F","船社","船名","指示元表示用","備考１","指示先データ","備考２","SZ","タイプ","高さ","材質","CYカット日","空コン搬出先","コード","担当","コード","担当")
'Chang 20050303 END

'2009/02/25 Add-S G.Ariola
dim ctr
for ctr = 1 to 3
Session(CSTR("Key" & ctr))
if Session(CSTR("Key" & ctr)) <> "" then
	Select Case Session(CSTR("Key" & ctr))
		Case "0" '入力日
			DtTbl(0)(0) = DtTbl(0)(0) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "1" '指示元 − コード
			DtTbl(0)(20) = DtTbl(0)(20) & getImage(Session(CSTR("KeySort" & ctr)))		
		Case "2" 'ブッキング番号
			DtTbl(0)(2) = DtTbl(0)(2) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "3" 'ピック済
			DtTbl(0)(3) = DtTbl(0)(3) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "4" 'SZ
			DtTbl(0)(14) = DtTbl(0)(14) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "5" 'タイプ
			DtTbl(0)(15) = DtTbl(0)(15) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "6" '高さ
			DtTbl(0)(16) = DtTbl(0)(16) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "7" '材質
			DtTbl(0)(17) = DtTbl(0)(17) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "8" '船社
			DtTbl(0)(8) = DtTbl(0)(8) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "9" '船名
			DtTbl(0)(9) = DtTbl(0)(9) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "10" 'CYカット日
			DtTbl(0)(18) = DtTbl(0)(18) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "11" '空コン搬出先
			DtTbl(0)(19) = DtTbl(0)(19) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "12" '指示先 − コード
			DtTbl(0)(4) = DtTbl(0)(4) & getImage(Session(CSTR("KeySort" & ctr)))
'		Case "13" '指示先 − 担当
'			DtTbl(0)(23) = DtTbl(0)(23) & getImage(Session(CSTR("KeySort" & ctr)))		
		Case "13" '指示先回答
			DtTbl(0)(5) = DtTbl(0)(5) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "14" '備考１
			DtTbl(0)(11) = DtTbl(0)(11) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "15" '備考２
			DtTbl(0)(13) = DtTbl(0)(13) & getImage(Session(CSTR("KeySort" & ctr)))
		Case "16" '指示元 − 担当
			DtTbl(0)(21) = DtTbl(0)(21) & getImage(Session(CSTR("KeySort" & ctr)))
	  End Select
end if	  
next
'2009/02/25 Add-E G.Ariola

  DtTbl(0)(6)=0
'3th Add Start
  Dim DelStr,DelTarget
  DelStr=Array("","未","No")
  DelTarget=Array(0,5,5)
'3th Add End

  'エラートラップ開始
    on error resume next
  'DB接続
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS

  '対象件数取得
    StrSQL = "SELECT count(SPB.BookNo) AS num FROM BookingAssign AS SPB "&_
             "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_
             strWrer
'3th Change SPBookInfo -> BookingAssign
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "2","b301","01","空搬出：一覧データ取得","101","SQL:<BR>"&strSQL
      Exit Function
    end if
    Num = ObjRS("num")
    ObjRS.close
    ReDim Preserve DtTbl(Num)
'3th ADD Start
    If Num>0 Then
'ADD 20050228 Fro survive ViewBookAssing ViewTable By SEIKO N.Oosige
'DEL 20100420 Start 
'      StrSQL = "IF (EXISTS( select * from ViewBookAssing ) ) BEGIN DROP VIEW ViewBookAssing END "
'      ObjConn.Execute(StrSQL)
'      if err <> 0 then
'        err.Clear
'      end if
'      
'      StrSQL = "BEGIN TRAN TRAN1 "
'      ObjConn.Execute(StrSQL)
'      if err <> 0 then
'        Set ObjRS = Nothing
'        jampErrerP "2","b301","01","空搬出：一覧データ取得","101","SQL:<BR>"&strSQL
'      end if
'ADD 20050228 End
'DEL 20100420 END 
'DEL 20100414 Start      
'     StrSQL = "CREATE VIEW ViewBookAssing AS SELECT Max(InputDate) AS MAXDATE,BookNo "&_
'               "FROM BookingAssign GROUP BY BookNo,Process "&_
'               "HAVING Process='R'"
''CW-319 ADD HAVING Process='R'
'      ObjConn.Execute(StrSQL)
'      if err <> 0 then
'        Set ObjRS = Nothing
'        jampErrerP "2","b301","01","空搬出：一覧データ取得","101","SQL:<BR>"&strSQL
'      end if
'3th ADD END
'DEL 20100414 END

  'データ取得
'CW-003    StrSQL = "SELECT BookNo, SenderCode, InputDate, TruckerCode, TruckerFlag "&_
'CW-003             "FROM SPBookInfo "&_
'CW-003             "WHERE (SenderCode='"& USER &"' OR TruckerCode='"& COMPcd &"') AND Process='R' "&_
'CW-003             strWrer &_
'CW-003             "ORDER BY InputDate ASC"
'CW-012    StrSQL = "SELECT Pickup.Qty, SPB.BookNo, SPB.SenderCode, SPB.InputDate, SPB.TruckerCode, SPB.TruckerFlag "&_
'CW-012             "FROM SPBookInfo AS SPB LEFT JOIN Pickup ON SPB.BookNo = Pickup.BookNo "&_
'CW-012             "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_
'CW-012              strWrer &_
'CW-012             "ORDER BY SPB.InputDate"
'    StrSQL = "SELECT MAX(Pic.Qty), SPB.BookNo, SPB.SenderCode, SPB.InputDate, SPB.TruckerCode, SPB.TruckerFlag, "&_
'             "mV.ShipLine, mV.FullName,mU.HeadCompanyCode, mU.UserType "&_
'             "FROM (((SPBookInfo AS SPB LEFT JOIN ExportCont AS EXC ON SPB.BookNo = EXC.BookNo) "&_
'             "LEFT JOIN Pickup AS Pic ON (EXC.BookNo = Pic.BookNo) AND (EXC.VoyCtrl = Pic.VoyCtrl) "&_
'             "AND (EXC.VslCode = Pic.VslCode)) "&_
'             "LEFT JOIN mVessel AS mV ON EXC.VslCode = mV.VslCode) "&_
'             "LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_
'             "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_
'              strWrer &_
'             "ORDER BY SPB.InputDate "
'CW-051 ADD This Line "mU.HeadCompanyCode, mU.UserType "&_
'CW-051 ADD This Line "INNER JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_


'2006/03/06 h.matsuda mod-s
'      StrSQL = "SELECT SPB.BookNo, SPB.SenderCode, SPB.InputDate, SPB.TruckerCode, SPB.TruckerFlag, "&_
'               "SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1, "&_
'               "SPB.Comment1,SPB.Comment2, mU.HeadCompanyCode, mU.UserType "&_
'               "FROM BookingAssign AS SPB LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_
'               "LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
'               "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_
'                strWrer &_
'               "ORDER BY VBA.MAXDATE DESC,SPB.InputDate DESC, SPB.BookNo ASC"
'      StrSQL = "SELECT SPB.BookNo, SPB.SenderCode, SPB.InputDate, SPB.TruckerCode, SPB.TruckerFlag, "&_
'               "SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1, "&_
'               "SPB.Comment1,SPB.Comment2, mU.HeadCompanyCode, mU.UserType "&_
'               ",SPB.ShipLine "&_
'               "FROM BookingAssign AS SPB LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_
'               "LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
'               "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_
'                strWrer &_
'               "ORDER BY VBA.MAXDATE DESC,SPB.InputDate DESC, SPB.BookNo ASC"
'2006/03/06 h.matsuda mod-s

'2009/02/25 Del-S G.Ariola  
'      StrSQL = "SELECT SPB.BookNo, SPB.SenderCode, SPB.InputDate, SPB.TruckerCode, SPB.TruckerFlag, "&_
'               "SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1, "&_
'               "SPB.Comment1,SPB.Comment2, mU.HeadCompanyCode, mU.UserType "&_
'               ",SPB.ShipLine "&_
'               "FROM BookingAssign AS SPB LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_
'               "LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
'               "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "&_
'                strWrer &_
'               "ORDER BY VBA.MAXDATE DESC,SPB.InputDate DESC, SPB.BookNo ASC"
'2009/02/25 Del-E G.Ariola  

      StrSQL = "SELECT T.* FROM (SELECT Distinct SPB.BookNo, mV.FullName, SPB.SenderCode, SPB.InputDate, SPB.TruckerCode, SPB.TruckerFlag, "&_
               "SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,SPB.ContMaterial1, "&_
			   
			   "(CASE "&_
			   "WHEN mU.UserType = '5' THEN mU.HeadCompanyCode "&_
			   "ELSE SPB.SenderCode END) as Code1, "&_
			
			   "(CASE "&_
			   "WHEN mU.UserType = '5' THEN mU.TTName "&_
			   "ELSE mU.TTName END) as TruckerName, "&_
			   "SPB.TruckerSubName as Name1, "&_			   
			   "(IsNull(CASE (SELECT count(distinct P.PickPlace) as Picks FROM Pickup AS P "&_
			   "WHERE BOK.VslCode = P.VslCode AND BOK.VoyCtrl = P.VoyCtrl AND BOK.BookNo = P.BookNo "&_
			   "Group BY P.BookNo) "&_
			   "WHEN '1' THEN Pic.PickPlace "&_
			   "ELSE '複数' END ,'')) PickPlace,  "&_

			   "(SELECT sum(ISDATE(EXC.EmpDelTime)) AS numC "&_
			   "FROM ExportCont AS EXC  "&_
			   "LEFT JOIN Container AS Con ON EXC.ContNo=Con.ContNo AND "&_
			   "EXC.VoyCtrl=Con.VoyCtrl AND EXC.VslCode=Con.VslCode "&_
			   "WHERE EXC.BookNo=SPB.BookNo) as NumCount, "&_
			
               "SPB.Comment1,SPB.Comment2, mU.HeadCompanyCode, mU.UserType "&_
               ",SPB.ShipLine,VSLS.CYCut "&_
               "FROM BookingAssign AS SPB LEFT JOIN mUsers AS mU ON SPB.SenderCode = mU.UserCode "&_
               "LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
			   
			   "LEFT JOIN ExportCont AS EXC ON EXC.bookno=SPB.bookno "&_
			   "left join (select a.bookno bookno ,b.vslcode vslcode , b.voyctrl voyctrl , "&_
			   "isnull(a.shipline,b.shipline) shipline "&_
			   "from bookingassign A left join booking b on a.bookno=b.bookno "&_
			   "WHERE (A.SenderCode='"& USER &"' OR A.TruckerCode='"& COMPcd &"') AND A.Process='R' "&_
			   strWrer2 &_
			   ") as BOK on exc.bookno=BOK.bookno and exc.vslcode=BOK.vslcode and exc.voyctrl=BOK.voyctrl "&_
			   
			   "LEFT JOIN VslSchedule AS VSLS ON BOK.VoyCtrl = VSLS.VoyCtrl AND BOK.VslCode = VSLS.VslCode "&_				   
			   "LEFT JOIN mVessel AS mV ON BOK.VslCode = mV.VslCode "&_
			   		   
			   "LEFT JOIN Pickup AS Pic ON BOK.VoyCtrl = Pic.VoyCtrl AND BOK.BookNo = Pic.BookNo AND BOK.VslCode = Pic.VslCode "&_
               "WHERE (SPB.SenderCode='"& USER &"' OR SPB.TruckerCode='"& COMPcd &"') AND SPB.Process='R' "& strWrer &") AS T "&_
             	strOrder

'"LEFT JOIN Booking AS BOK ON SPB.BookNo = BOK.BookNo "&_
'response.write StrSQL
'response.End
'C-002 ADD This Item : SPB.Comment
'20030910 chage "ORDER BY SPB.InputDate ASC"->"ORDER BY SPB.InputDate DESC"
'3th Change SPBookInfo -> BookingAssign
'3th Change Comment -> Comment1,Comment2
'3th ADD SPB.ContSize1,SPB.ContType1,SPB.ContHeight1,
'3th ADD Line LEFT JOIN ViewBookAssing AS VBA ON SPB.BookNO=VBA.BookNo "&_
'3th ADD VBA.MAXDATE DESC and SPB.BookNo ASC
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "2","b301","01","空搬出：一覧データ取得","102","SQL:<BR>"&strSQL
        Exit Function
      end if
      dim tmpBookingNo
      tmpBookingNo=""
      i=1

      Do Until ObjRS.EOF
'CW-002      DtTbl(i)=Array("","","","","","")
'CW-003      DtTbl(i)=Array("","","","","","","")
'C-002      DtTbl(i)=Array("","","","","","","","","","","")
'3th      DtTbl(i)=Array("","","","","","","","","","","","")
        'DtTbl(i)=Array("","","","","","","","","","","","","","","","","","")
		DtTbl(i)=Array("","","","","","","","","","","","","","","","","","","","","","")
        DtTbl(i)(0)=Mid(ObjRS("InPutDate"),3,8)
        DtTbl(i)(1)=Trim(ObjRS("SenderCode"))
'        If ObjRS("UserType") = "5" Then			'CW-051
'          DtTbl(i)(10)=Trim(ObjRS("HeadCompanyCode"))	'CW-051
'        Else						'CW-051
'          DtTbl(i)(10)=Trim(ObjRS("SenderCode"))	'CW-051
'        End If						'CW-051
		DtTbl(i)(10)=Trim(ObjRS("Code1"))
		DtTbl(i)(20)=Trim(ObjRS("Name1"))
		DtTbl(i)(21)=Trim(ObjRS("TruckerName"))
        DtTbl(i)(2)=Trim(ObjRS("BookNo"))
        DtTbl(i)(4)=Trim(ObjRS("TruckerCode"))
        DtTbl(i)(12)=DtTbl(i)(4)
        DtTbl(i)(6)=ObjRS("TruckerFlag")		'CW-002
        If DtTbl(i)(1) <> USER AND DtTbl(i)(6)=0 Then
          DtTbl(i)(6)=3
        End If
'        DtTbl(i)(8)=Trim(ObjRS("ShipLine"))
'        DtTbl(i)(9)=Left(ObjRS("FullName"),12)
        DtTbl(i)(8)=Trim(ObjRS("ShipLine"))
        DtTbl(i)(9)=Trim(ObjRS("FullName"))
        DtTbl(i)(14)=Trim(ObjRS("ContSize1"))
        DtTbl(i)(15)=Trim(ObjRS("ContType1"))
        DtTbl(i)(16)=Trim(ObjRS("ContHeight1"))
        DtTbl(i)(17)=Trim(ObjRS("ContMaterial1"))
		DtTbl(i)(18)=Trim(Mid(ObjRS("CYCut"),3,8))
		DtTbl(i)(19)=Trim(ObjRS("PickPlace"))
        If DtTbl(i)(1) = USER AND DtTbl(i)(4)<>COMPcd AND DtTbl(i)(4)<>""  Then
        '指示先照会済みフラグ
          If ObjRS("TruckerFlag")=0 Then
            DtTbl(i)(5) = "未"
          ElseIf ObjRS("TruckerFlag")=1 Then
            DtTbl(i)(5) = "Yes"
          Else
            DtTbl(i)(5) = "No"
          End If
          DtTbl(i)(6) = 0
        Else
          DtTbl(i)(4) = "　"
          DtTbl(i)(5) = "　"
        End If
      
'3th      DtTbl(0)(6)=DtTbl(0)(6)+DtTbl(i)(6)
'      DtTbl(i)(7)=Trim(ObjRS("Qty"))		'CW-003
'      If IsNull(DtTbl(i)(7)) Then		'CW-003
'        DtTbl(i)(7)=0				'CW-003
'      End If					'CW-003
        'DtTbl(i)(7)=0
		DtTbl(i)(7)=Trim(ObjRS("NumCount"))
        'DtTbl(i)(3)=0
		DtTbl(i)(3)=Trim(ObjRS("NumCount"))
'3th      DtTbl(i)(11)=ObjRS("Comment")	'C-002
        DtTbl(i)(11)=ObjRS("Comment1")
        DtTbl(i)(13)=ObjRS("Comment2")
        If DtTbl(i)(2)<>DtTbl(i-1)(2) Then
          tmpBookingNo=tmpBookingNo&",'"&DtTbl(i)(2)&"'"
        End If
'3th Add Start
        If DelType=0 OR DtTbl(i)(DelTarget(DelType)) = DelStr(DelType) Then
          DtTbl(0)(6) = DtTbl(0)(6) + DtTbl(i)(6)
          i=i+1
        Else
          Num=Num-1
        End If
'      i=i+1
'3th Add End
        ObjRS.MoveNext
      Loop
      ObjRS.close
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "2","b301","01","空搬出：一覧データ取得","200",""
        Exit Function
      end if
'3th ADD Start
'Change 20050228 Fro survive ViewBookAssing ViewTable By SEIKO N.Oosige
'      StrSQL = "DROP VIEW ViewBookAssing"
'DEL 20100420 Start
 '     StrSQL = "COMMIT TRAN TRAN1 "
''Change 20050228 End
'      ObjConn.Execute(StrSQL)
'      if err <> 0 then
'        Set ObjRS = Nothing
'        jampErrerP "2","b301","01","空搬出：一覧データ取得","101","SQL:<BR>"&strSQL
'      end if
'DEL 20100420 END
'3th ADD END
'2009/02/25 Del-S G.Ariola  
      'ピック済み本数取得
'      If tmpBookingNo <> "" Then
''2006/03/06 mod-s h.matsuda(SQL文を再構築)
''        StrSQL = "SELECT  EXC.BookNo,sum(ISDATE(EXC.EmpDelTime)) AS numC, mV.FullName, mV.ShipLine,ISNULL(Pic.Qty,0) AS Qty "&_
''                 "FROM ExportCont AS EXC LEFT JOIN mVessel AS mV ON EXC.VslCode = mV.VslCode "&_
''                 "LEFT JOIN Container AS Con ON EXC.ContNo=Con.ContNo AND EXC.VoyCtrl=Con.VoyCtrl AND EXC.VslCode=Con.VslCode "&_
''                 "LEFT JOIN Pickup AS Pic ON (EXC.BookNo = Pic.BookNo) AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.PickPlace=Pic.PickPlace) "&_
''                 "AND (EXC.VslCode = Pic.VslCode) AND (CON.ContSize=Pic.ContSize) AND (CON.ContType=Pic.ContType) AND (CON.ContHeight=Pic.ContHeight) "&_
''                 "WHERE EXC.BookNo IN("& Mid(tmpBookingNo,2) &") "&_
''                 "Group By EXC.BookNo, mV.FullName, mV.ShipLine,Pic.Qty "&_
''                 "Order By EXC.BookNo ASC "
'        strSQL = "          SELECT  EXC.BookNo                                                  "
'        strSQL = strSQL & " ,sum(ISDATE(EXC.EmpDelTime)) AS numC, mV.FullName                   "
'        strSQL = strSQL & " ,bkg.shipline                                                       "
'        strSQL = strSQL & " ,ISNULL(Pic.Qty,0) AS Qty                                           "
'        strSQL = strSQL & " FROM ExportCont AS EXC                                              "
'        strSQL = strSQL & " LEFT JOIN mVessel AS mV ON EXC.VslCode = mV.VslCode                 "
'        strSQL = strSQL & " LEFT JOIN Container AS Con ON EXC.ContNo=Con.ContNo                 "
'        strSQL = strSQL & " AND EXC.VoyCtrl=Con.VoyCtrl AND EXC.VslCode=Con.VslCode             "
'        strSQL = strSQL & " LEFT JOIN Pickup AS Pic ON (EXC.BookNo = Pic.BookNo)                "
'        strSQL = strSQL & " AND (EXC.VoyCtrl = Pic.VoyCtrl) AND (EXC.PickPlace=Pic.PickPlace)   "
'        strSQL = strSQL & " AND (EXC.VslCode = Pic.VslCode)                                     "
'        strSQL = strSQL & " AND (CON.ContSize=Pic.ContSize)                                     "
'        strSQL = strSQL & " AND (CON.ContType=Pic.ContType)                                     "
'        strSQL = strSQL & " AND (CON.ContHeight=Pic.ContHeight)                                 "
'        strSQL = strSQL & " left join (select a.bookno bookno ,b.vslcode vslcode ,              "
'        strSQL = strSQL & " b.voyctrl voyctrl ,isnull(a.shipline,b.shipline) shipline           "
'        strSQL = strSQL & " from bookingassign A left join booking b                            "
'        strSQL = strSQL & " on a.bookno=b.bookno where a.sendercode='" & USER & "') as bkg      "
'        strSQL = strSQL & " on exc.bookno=bkg.bookno                                            "
'        strSQL = strSQL & " and exc.vslcode=bkg.vslcode and exc.voyctrl=bkg.voyctrl             "
'        strSQL = strSQL & " WHERE EXC.BookNo IN(" & Mid(tmpBookingNo, 2) & ")                      "
'        strSQL = strSQL & " Group By EXC.BookNo, mV.FullName, bkg.shipline,Pic.Qty              "
'        strSQL = strSQL & " Order By EXC.BookNo ASC                                             "
''2006/03/06 add-e h.matsuda(SQL文を再構築)
''response.Write(StrSQL)
''response.End()
'        ObjRS.Open strSQL, ObjConn
'        If Err <> 0 Then
'          DisConnDBH ObjConn, ObjRS 'DB切断
'            jampErrerP "2", "b301", "01", "空搬出：一覧データ取得", "102", "SQL:<BR>" & strSQL
'          Exit Function
'        End If
'        ReDim tmpBookingNo(Num)
'        tmpBookingNo(0) = Array("", 0, "", "", 0)
'        i = 1
'        tmpBookingNo(1) = Array("", 0, "", "", 0)
'        Do Until ObjRS.EOF
'          If tmpBookingNo(i - 1)(0) = Trim(ObjRS("BookNo")) Then
'            tmpBookingNo(i - 1)(1) = tmpBookingNo(i - 1)(1) + ObjRS("numC")
'            tmpBookingNo(i - 1)(4) = tmpBookingNo(i - 1)(4) + ObjRS("Qty")
'          Else
'            tmpBookingNo(i)(0) = Trim(ObjRS("BookNo"))
'            tmpBookingNo(i)(1) = ObjRS("numC")
'            tmpBookingNo(i)(2) = Trim(ObjRS("ShipLine"))
'            tmpBookingNo(i)(3) = Trim(ObjRS("FullName"))
'            tmpBookingNo(i)(4) = ObjRS("Qty")
'            i = i + 1
'            tmpBookingNo(i) = Array("", 0, "", "", 0)
'          End If
'          ObjRS.MoveNext
'        Loop
'        tmpBookingNo(0)(1) = i - 1
'        ObjRS.Close
'        For i = 1 To Num
'          For j = 1 To tmpBookingNo(0)(1)
'            If DtTbl(i)(2) = tmpBookingNo(j)(0) Then
'              '2009/02/25 Del-S G.Ariola
'              'DtTbl(i)(3) = tmpBookingNo(j)(1)
'              'DtTbl(i)(8) = tmpBookingNo(j)(2)
'              'DtTbl(i)(9) = tmpBookingNo(j)(3)
'              '2009/02/25 Del-E G.Ariola
'              If tmpBookingNo(j)(1) = tmpBookingNo(j)(4) Then
'                DtTbl(i)(7) = tmpBookingNo(j)(4)
'              End If
'            End If
'          Next
'        Next
'      End If
''      For i=1 To Num
''        StrSQL = "SELECT Count(BookNo) AS numC FROM ExportCont "&_
''                 "WHERE BookNo='"& DtTbl(i)(2) &"' AND EmpDelTime IS NOT NULL"
''        ObjRS.Open StrSQL, ObjConn
''        if err <> 0 then
''          DisConnDBH ObjConn, ObjRS    'DB切断
''          jampErrerP "2","b301","01","空搬出：一覧データ取得","101","SQL:<BR>"&strSQL
''          Exit Function
''        end if
''        DtTbl(i)(3) = ObjRS("numC")
''CW-020      If DtTbl(i)(7)<>"0" AND DtTbl(i)(7)<>DtTbl(i)(3) Then  'CW-003
''        If DtTbl(i)(7)<>"0" AND CInt(DtTbl(i)(7))<>CInt(DtTbl(i)(3)) Then  'CW-020
''          DtTbl(i)(7)=0                        'CW-003
''        End If                         'CW-003
''        ObjRS.close
''      Next

'2009/02/25 Del-E G.Ariola  
  End If        'If Num>0    3th ADD
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
function vew(){
<%If Num<>0 Then%> 
// var obj1=document.getElementById("HDIV");
// var obj2=document.getElementById("VDIV");
var obj3=document.getElementById("BDIV");
/* if((document.body.offsetWidth-10)<1243){
    obj1.style.width=document.body.offsetWidth-10;
	obj1.style.overflow="auto";
 }
  else
 {
 	 obj1.style.width=document.body.offsetWidth-10;
	 obj1.style.overflowX="auto";
 }
 obj2.style.height=document.body.offsetHeight-100;
*/ 
 if((document.body.offsetWidth-10)  < 1243){
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
//function GoRenew(bookNo,compF,SijiM,SijiC,sShipLine){// mod-s h.matsuda 2006/03/06
function GoRenew(bookNo,compF,SijiM,SijiC,sShipLine){// mod-s h.matsuda 2006/03/06
  Fname=document.dmo210F;
  Fname.BookNo.value=bookNo;
  Fname.CompF.value=compF;
  Fname.COMPcd0.value=SijiM;
  Fname.COMPcd1.value=SijiC;
// 2006/03/06 mod h.matsua  
  Fname.ShipLine.value=sShipLine;
  Fname.action="./dmi312.asp";
//  Fname.action="./dmi215.asp";
// 2006/03/06 mod h.matsua  
  newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes,scrollbars=yes");
  Fname.target="ReEntry";
  Fname.submit();
}
//検索
function SerchC(SortFlag,Kye){
  Fname=document.dmo210F;
  Fname.SortFlag.value=SortFlag;
  Fname.SortKye.value=Kye;
  Fname.target="_self";
  Fname.action="./dmo210L.asp";
  Fname.submit();
}
//照会済
function GoSyokaizumi(){
  target=document.dmo210F;
  if(target.DataNum.value>0){
    flag = confirm('未回答の回答を「Yes」にしますか？');
    if(flag==true){
      target.SortFlag.value=8;
      target.target="_self";
      target.action="./dmo210L.asp";
      len=target.elements.length;
      for(i=0;i<len;i++){
        target.elements[i].disabled=false;
      }
      target.submit();
    }
  }
}
//CSV		ADD C-001
function GoCSV(){
  target=document.dmo210F;
  len=target.elements.length;
  for(i=0;i<len;i++){
    target.elements[i].disabled=false;
  }
//  target.target="Bottom";
  target.action="./dmo280.asp";
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
<style>
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
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setTimeout('showContent()', 500);vew();" onResize="vew()">
<div id="content" style="display:none;">
<!-------------空搬出情報一覧画面List--------------------------->
<Form name="dmo210F" method="POST">
一覧に表示される属性（サイズ等）は登録時に先頭に入力されたもののみです。各データの詳細画面では全件表示されます。<BR>
<div id="content" style="display:none;">
<!--<DIV style=" width:1043; height:41;"> --><!--2009/02/25 G.Ariola -->
<div id="BDIV">
<TABLE border="1" cellPadding="2" cellSpacing="0">
<%If Num<>0 Then%>
  <%If DtTbl(0)(6) = 0 Then %>
  <thead>
  <tr>
    <TH class="hlist" nowrap><%=DtTbl(0)(0)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(1)%><BR><%=DtTbl(0)(20)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(2)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(14)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(9)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(11)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(13)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(21)%>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>'>
    </TH>
  </TR>
  <!--2009/02/25 Add-S G.Ariola test1 -->
  	<!--<TR class=bga>
	<TH width="100"><'%=DtTbl(0)(20)%></TH>
	<TH width="100"><'%=DtTbl(0)(21)%></TH> -->
	<!--<TH width="60"><'%=DtTbl(0)(22)%></TH>
	<TH width="60"><'%=DtTbl(0)(23)%></TH> -->
	<!--</TH></TR> -->
	<!--2009/02/25 Add-E G.Ariola -->
<!--</TABLE> --> <!--2009/02/25 G.Ariola -->
<!--</DIV> --><!--2009/02/25 G.Ariola -->
</THEAD>
<%'If Num>10 Then%>  <!--<DIV id="VDIV" style=" width:999; height:242; overflow-y:scroll;"> --><!--2009/02/25 G.Ariola -->
  <%'else%><!--<DIV id="VDIV" style=" width:983; height:242;"> --><!--2009/02/25 G.Ariola -->
 <%'end if%> <!--2009/02/25 G.Ariola -->
<!--<TABLE border="1" cellPadding="2" cellSpacing="0" cols="<%=Num+20%>"> --><!--2009/02/25 G.Ariola -->  
<tbody>    
    <% For j=1 to Num %>
  <TR class=bgw>
    <TD nowrap><%=DtTbl(j)(0)%><BR></TD><TD nowrap><%=DtTbl(j)(10)%><!--<INPUT value="" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 > --><BR></TD>

<%'Mod-s 2006/03/06 h.matsuda--->%>
<!--    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(2)%>','<%=DtTbl(j)(7)%>','<%=DtTbl(j)(1)%>','<%=DtTbl(j)(12)%>');">-->
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(2)%>','<%=DtTbl(j)(7)%>','<%=DtTbl(j)(1)%>','<%=DtTbl(j)(12)%>','<%=DtTbl(j)(8)%>');">
<%'Mod-e 2006/03/06 h.matsuda--->%>

        <%=DtTbl(j)(2)%></A></TD><TD nowrap><%=DtTbl(j)(3)%><BR></TD><TD nowrap><%=DtTbl(j)(14)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=DtTbl(j)(16)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(8)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(9),12)%><BR></TD><TD nowrap><%=DtTbl(j)(18)%><BR></TD><TD nowrap><INPUT value="<%=DtTbl(j)(19)%>" type=text class=chrReadOnly size="21" readonly=TRUE tabindex = -1 ><BR></TD><TD nowrap><%=DtTbl(j)(4)%><BR></TD><!--<TD width="57"><INPUT value="<%=DtTbl(j)(21)%>" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 ><BR></TD> -->
    <TD nowrap><%=DtTbl(j)(5)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(11),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(13),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(20),8)%><BR>
	
    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>'>
    </TD>
  </TR>
    <% Next %>
<!--</TABLE> -->	<!--2009/02/25 G.Ariola -->
</tbody>
<%'If Num>10 Then%><!--</DIV> --><%'end if%>	<!--2009/02/25 G.Ariola -->			
  <% Else %>
  <thead>  
     <tr >
    <TH class="hlist" nowrap><%=DtTbl(0)(0)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(1)%><BR><%=DtTbl(0)(20)%></TH><TH class="hlist" nowrap>指示元<BR>へ回答</TH><TH class="hlist" nowrap><%=DtTbl(0)(2)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(3)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(14)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(15)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(16)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(17)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(8)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(9)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(18)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(19)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(4)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(5)%></TH>
    <TH class="hlist" nowrap><%=DtTbl(0)(11)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(13)%></TH><TH class="hlist" nowrap><%=DtTbl(0)(21)%>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>,<%=DtTbl(0)(15)%>,<%=DtTbl(0)(16)%>,<%=DtTbl(0)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>'>
    </TH>
  </TR>
  <!--2009/02/25 Add-S G.Ariola -->
  <!--<TR class=bga>
	<TH width="100"><'%=DtTbl(0)(20)%></TH>
	<TH width="100"><'%=DtTbl(0)(21)%></TH> -->
	<!--<TH width="60"><'%=DtTbl(0)(22)%></TH>
	<TH width="60"><'%=DtTbl(0)(23)%></TH> -->
	<!--</TH></TR> -->
  <!--2009/02/25 Add-S G.Ariola -->
  <!--</TABLE> --> <!--2009/02/25 G.Ariola -->
  <!--</DIV> --><!--2009/02/25 G.Ariola -->
  </thead>  
<%'If Num>10 Then%>  <!--<DIV id="VDIV" style=" width:999; height:242; overflow-y:scroll;"> --><!--2009/02/25 G.Ariola -->
  <%'else%><!--<DIV id="VDIV" style=" width:986; height:242;"> --><!--2009/02/25 G.Ariola -->
 <%'end if%> <!--2009/02/25 G.Ariola -->
  <!--<TABLE border="1" cellPadding="2" cellSpacing="0" cols="<%=Num+20%>">   -->
  <tbody>
    <% For j=1 to Num %>
  <TR class=bgw>
    <TD nowrap><%=DtTbl(j)(0)%></TD><TD nowrap><%=DtTbl(j)(10)%><BR><!--<INPUT value="" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 > --></TD>
    <TD nowrap><%=ResA(DtTbl(j)(6))%></TD>

<%'Mod-s 2006/03/06 h.matsuda--->%>
<!--    <TD width="60"><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(2)%>','<%=DtTbl(j)(7)%>','<%=DtTbl(j)(1)%>','<%=DtTbl(j)(12)%>');">-->
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(2)%>','<%=DtTbl(j)(7)%>','<%=DtTbl(j)(1)%>','<%=DtTbl(j)(12)%>','<%=DtTbl(j)(8)%>');">
<%'Mod-e 2006/03/06 h.matsuda --------------------------->%>

        <%=DtTbl(j)(2)%></A></TD><TD nowrap><%=DtTbl(j)(3)%><BR></TD><TD nowrap><%=DtTbl(j)(14)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(15)%><BR></TD><TD nowrap><%=DtTbl(j)(16)%><BR></TD><TD nowrap><%=DtTbl(j)(17)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(8)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(9),12)%><BR></TD><TD nowrap><%=DtTbl(j)(18)%><BR></TD><TD nowrap><INPUT value="<%=DtTbl(j)(19)%>" type=text class=chrReadOnly size="21" readonly=TRUE tabindex = -1 ><BR></TD>
    <TD nowrap><%=DtTbl(j)(4)%><BR></TD><!--<TD width="57"><INPUT value="<%=DtTbl(j)(21)%>" type=text class=chrReadOnly size="9" readonly=TRUE tabindex = -1 ><BR></TD> --><TD nowrap><%=DtTbl(j)(5)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(11),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(13),10)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(20),8)%><BR>
    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>,<%=DtTbl(j)(15)%>,<%=DtTbl(j)(16)%>,<%=DtTbl(j)(17)%>,<%=DtTbl(0)(18)%>,<%=DtTbl(0)(19)%>'>
    </TD>
  </TR>
    <% Next %>
 <!--</TABLE> -->	<!--2009/02/25 G.Ariola -->
 </tbody>
<%'If Num>10 Then%><!--  </DIV> --><%'end if%>	<!--2009/02/25 G.Ariola -->		
  <% End If %>
<% Else %>
  <TR class=bgw><TD nowrap>作業案件はありません</TD></TR>
<% End If %>
</TABLE>
</DIV>
<%'3th del Set_Data Num,DtTbl %>
  <INPUT type=hidden name=DataNum value="<%=Num%>">
  <INPUT type=hidden name=SortFlag value="<%=SortFlag%>" >
  <INPUT type=hidden name=SortKye value="<%=SortKye%>" >
  <INPUT type=hidden name=BookNo value="" >
  <INPUT type=hidden name=CompF value="" >
  <INPUT type=hidden name=COMPcd0 value="" >
  <INPUT type=hidden name=COMPcd1 value="" >
  <INPUT type=hidden name=Mord value="1" >
  <INPUT type=hidden name=strWhere value="<%=strWrer%>" disabled>
<%'Mod-s 2006/03/06 h.matsuda--->%>
	  <INPUT type=hidden name="ShoriMode" value="EMoutUpd">
	  <INPUT type=hidden name="ShipLine" value="">
<%'Mod-e 2006/03/06 h.matsuda --------------------------->%>
</Form>
</div>
<!-------------画面終わり--------------------------->
</BODY></HTML>
