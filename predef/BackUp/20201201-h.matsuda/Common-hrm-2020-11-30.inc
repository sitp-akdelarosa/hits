<%
'
'SQLサーバへの接続設定
'dim Host,DBName,ID,PassW
'Host  ="192.168.16.14"		'DBサーバアドレス
'DbName="wavedb"			'DB名称
'ID    ="wave"			'DBログインID
'PassW ="wave"			'DBログインパスワード

'
'SQLサーバへの接続
'Input :DB接続Object、レコードセットObject
'
'Function ConnDBH(ObjConn, ObjRS)
' Set ObjConn = Server.CreateObject("ADODB.Connection")
'  ObjConn.Open "Provider=SQLOLEDB;Data Source=" & Host & ";" &_
'               "Database=" & DbName & ";User ID=" & ID & ";Password=" & PassW & ";"
'  Set ObjRS = Server.CreateObject("ADODB.Recordset")
'  ObjConn.BeginTrans
'End Function

Function ConnDBH(ObjConn,ObjRS)
  Set ObjConn = Server.CreateObject("ADODB.Connection")
  'ObjConn.Open "wavedb","wave","wave"
  'ObjConn.Open "wavedb","sa","sa"
  ObjConn.Open "Provider=SQLOLEDB; Data Source = SI-KURT-LAPTOP\SQLEXPRESS; Initial Catalog = OPE3_KAM; User Id = sa; Password=P@55w0rd"
  
  'ObjConn.ConnectionTimeout = 10              'SQL接続ﾀｲﾑｱｳﾄ時間設定   Add 20080206
  'ObjConn.CommandTimeout    = 120             'SQL実行ﾀｲﾑｱｳﾄ時間設定   Add 20080206
  
  ObjConn.CursorLocation = 3	'''3:adUseClient	'''Added 20040227

  Set ObjRS = Server.CreateObject("ADODB.Recordset")
  ObjConn.BeginTrans
End Function

'2016/11/08 H.Yoshikawa Add Start
'KACCS_OPEへのDB接続
Function ConnDBOpe(Ope,ObjConn,ObjRS)
  Set ObjConn = Server.CreateObject("ADODB.Connection")
  select case Ope
  	case "HKK"
  		ObjConn.Open "HKK-MAIN","KOUN","KO!201611UN"
  	case "KAM"
  		ObjConn.Open "KAM-MAIN","KAMI","KA!201611MI"
  	case "KTC"
  		ObjConn.Open "KTC-MAIN","GENEQ","GE!201611NEQ"
  	case "MLC"
  		ObjConn.Open "MLC-MAIN","MITU","MI!201611TU"
  	case "NEC"
  		ObjConn.Open "NEC-MAIN","NITTSU","NI!201611TTSU"
  	case "SOG"
  		ObjConn.Open "SOG-MAIN","SOGO","SO!201611GO"
  end select

  Set ObjRS = Server.CreateObject("ADODB.Recordset")
  ObjConn.BeginTrans
End Function
'2016/11/08 H.Yoshikawa Add End

'
'SQLサーバへの切断
'Input :DB接続Object、レコードセットObject
'
Function DisConnDBH(ObjConn,ObjRS)
  Set ObjRS = Nothing
  ObjConn.CommitTrans
  ObjConn.Close
  Set ObjConn = Nothing
End Function

'
'ユーザーＩＤが有効かどうかのチェック
'Input :なし
'
Function CheckLoginH()
    ' セッションの有効性をチェック
    Dim strUserID
    strUserID = Trim(Session.Contents("userid"))

    ' セッションが無効なとき
    If strUserID="" Then
      dim strtmp
      strtmp = "ログインしてから"&Session.Timeout&"分経過しました。"
      jampErrerP "0","b000","00","事前情報入力","000",strtmp
    End If
End Function
'
'アクセスログ出力
'Input :画面ID、画面名、作業ID、確認用オプションデータ
'
Function WriteLogH(iNum, strScrName, iWrkNum, strOption)
  'サーバ日時の取得
    dim DayTime,day
    getDayTime DayTime
  'ユーザタイプコード
    dim UtypeA
    UtypeA = Array("","荷主","海貨","船社","港運","陸運")
    dim strDate,logPath,ObjFSO,ObjTS

    strDate = DayTime(0) & DayTime(1) & DayTime(2)
'    logPath = Server.MapPath("../log/" & strDate & ".log")
    logPath = Server.MapPath("/log/" & strDate & ".log")
'    logPath = Server.MapPath("/20090316/log/" & strDate & ".log")

    Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set ObjTS = ObjFSO.OpenTextFile(logPath,8,True)
    ObjTS.Write DayTime(0)&"/"&DayTime(1)&"/"&DayTime(2)&" "&DayTime(3)&":"&DayTime(4)&":"&Right("0" & Second(Now), 2)
    ObjTS.Write "," & iNum
    ObjTS.Write "," & strScrName
    ObjTS.Write "," & iWrkNum
    ObjTS.Write "," & UtypeA(Session.Contents("UType")) & "," & Session.Contents("userid")
    ObjTS.Write "," & Request.ServerVariables("SCRIPT_NAME")
    ObjTS.Write "," & Request.ServerVariables("REMOTE_ADDR")
    If strOption<>"" Then
        ObjTS.WriteLine "," & strOption
    Else
        ObjTS.WriteLine 
    End If
    ObjTS.Close
    Set ObjTS = Nothing
    Set ObjFSO = Nothing

    '2009/03/05 Add-S Tanaka
     xLogReg iNum,iWrkNum,Session.Contents("userid")
    '2009/03/05 Add-E Tanaka

End Function

'2009/03/05 Add-S Tanaka
'ログをDBに登録
Function xLogReg(iNum,iWrkNum,strUserId)
	'iNum		[ I ]画面ID
	'iWrkNum	[ I ]作業ID

	Dim ObjConn, ObjRS, StrSQL,fs,ti,strTemp
	Dim iAddFlag
	Dim PageNum(),WkNum(),PageTitle(),SubTitle(),anyTmpTitle
	Dim strTitleFileName
	dim DayTime,day,strDate
	getDayTime DayTime

	strDate = DayTime(0) & DayTime(1) & DayTime(2)

	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	
	'DB接続
	ConnDBH ObjConn,ObjRS


	' ログタイトル取得
	strTitleFileName="../logweb.txt"
'	strTitleFileName="/20090316/logweb.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iAddFlag=0
	
	'ログタイトル分ループ
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		'ログタイトルに該当するもの場合
		IF TRIM(anyTmpTitle(0))=TRIM(CSTR(iNum)) and TRIM(anyTmpTitle(1))=TRIM(CSTR(iWrkNum)) then
			'ログインコードが該当のもの
			IF strUserId<>"90000" and strUserId<>"10000" and strUserId<>"MES" and strUserId<>"MEO" Then
				iAddFlag=1
				Exit Do
			ELSE
				Exit Function
			End If
		End If
	Loop
	ti.Close

	'登録可能なデータが存在する場合
	If iAddFlag=1 Then

		StrSQL = ""
		StrSQL = StrSQL & " SELECT * FROM xLog "
		StrSQL = StrSQL & " WHERE LogDate='" & strDate &"'"
		StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
		StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		
		ObjRS.Open StrSQL, ObjConn

		'レコードが存在する場合
		If ObjRS.EOF=False then
			StrSQL = ""
			StrSQL = StrSQL & " UPDATE xLog SET "
			StrSQL = StrSQL & " DataCount=DataCount+1 "
			StrSQL = StrSQL & " ,UpdtTime=GetDate() "
			StrSQL = StrSQL & " ,UpdtPgCd='COMMNINC' "
			StrSQL = StrSQL & " ,UpdtTmnl='" & Session.Contents("userid") & "'"
			StrSQL = StrSQL & " WHERE LogDate='" &strDate &"'"
			StrSQL = StrSQL & " AND PageNum='" & TRIM(CSTR(iNum)) &"'"
			StrSQL = StrSQL & " AND WkNum='" & TRIM(CSTR(iWrkNum)) &"'"
		ELSE
			StrSQL = ""
			StrSQL = StrSQL & " INSERT INTO xLog( "
			StrSQL = StrSQL & " LogDate, "
			StrSQL = StrSQL & " PageNum, "
			StrSQL = StrSQL & " WkNum, "
			StrSQL = StrSQL & " UpdtPgCd, "
			StrSQL = StrSQL & " UpdtTime, "
			StrSQL = StrSQL & " UpdtTmnl, "
			StrSQL = StrSQL & " DataCount "
			StrSQL = StrSQL & " )VALUES( "
			StrSQL = StrSQL & "'" & strDate & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iNum)) & "',"
			StrSQL = StrSQL & "'" & TRIM(CSTR(iWrkNum)) & "',"
			StrSQL = StrSQL & "'COMMNINC',"
			StrSQL = StrSQL & "GetDate(),"
			StrSQL = StrSQL & "'" & Session.Contents("userid") & "',"
			StrSQL = StrSQL & "1)"
		End If
		ObjConn.Execute(StrSQL)
		DisConnDBH ObjConn,ObjRS
	End If

End Function
'2009/03/05 Add-E Tanaka
'
'エラー画面への遷移
'Input :エラー画面フラグ、画面ID、作業ID、作業名、エラーコード、オプション
'
Function jampErrerP(WinFlag,dispId,wkID,wkName,errerCd,etc)
  on error goto 0		'エラートラップ解除
  Session.Contents("WinFlag")=WinFlag
  Session.Contents("dispId")=dispId
  Session.Contents("wkID")=wkID
  Session.Contents("wkName")=wkName
  Session.Contents("errerCd")=errerCd
  Session.Contents("etc")=etc
  Response.Redirect "./errer.asp"
End Function
'
'エラー画面への遷移
'Input :DB接続Object、レコードセットObject、エラー画面フラグ、画面ID、作業ID、作業名、エラーコード、オプション
'
Function jampErrerPDB(ObjConn,WinFlag,dispId,wkID,wkName,errerCd,etc)
'ロールバックとDB切断
  ObjConn.RollbackTrans
  ObjConn.Close
  Set ObjConn = Nothing
  on error goto 0		'エラートラップ解除
  Session.Contents("WinFlag")=WinFlag
  Session.Contents("dispId")=dispId
  Session.Contents("wkID")=wkID
  Session.Contents("wkName")=wkName
  Session.Contents("errerCd")=errerCd
  Session.Contents("etc")=etc
  Response.Redirect "./errer.asp"
End Function

'
'サーバ日時の取得
'Input :Array(5)=(YYYY,MM,DD,hh,mm)
'
Function getDayTime(DayTime)
  dim MyDate
  MyDate = Now
  DayTime = Array(Year(MyDate),Right("0"&Month(MyDate),2),Right(Date(),2),Right("0"&Hour(MyDate),2),Right("0"&Minute(MyDate),2))
End Function
'
'輸入コンテナテーブルのコンテナ存在チェック
'Input :DB接続Object、レコードセットObject、コンテナ／BL番号、作業フラグ、戻り値
'Output:ret(存在する=>true　存在しない=>false) 
Function checkImportCont(ObjConn, ObjRS,strNum, Flag, ret)
  dim StrSQL
  If Flag <> 4 Then
    StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE ContNo=" & strNum
  Else
    StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE BLNo=" & strNum
  End If
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b000","00","輸入コンテナテーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("num") <> 0 Then
    ret = true
  Else
    ret = false
  End If
  ObjRS.close
End Function
'
'CW-014 add strart
'
'輸入コンテナテーブルのコンテナ完了チェック
'Input :DB接続Object、レコードセットObject、コンテナ／BL番号、作業フラグ、戻り値
'Output:ret(未完了=>true　完了=>false) 
Function checkImportContComp(ObjConn, ObjRS,strNum, Flag, ret)
  dim StrSQL
  If Flag <> 4 Then
    StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE ContNo=" & strNum 
  Else
    StrSQL = "SELECT count(ContNo) AS num FROM ImportCont WHERE BLNo=" & strNum 
  End If
  StrSQL = StrSQL & " AND CYDelTime Is Null "
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "1","b000","00","輸入コンテナテーブル完了チェック","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("num") <> 0 Then
    ret = true
  Else
    ret = false
  End If
  ObjRS.close
End Function
'CW-014 add end
'
'IT共通テーブルの重複チェック
'Input :DB接続Object、レコードセットObject、コンテナ／BL番号、処理区分、作業フラグ、作業番号、戻り値
'Output:ret(重複なし=>true　重複有り=>false) 
'      :WkNo(作業番号)
Function checkComInfo(ObjConn, ObjRS, strNum, WFlag, Flag, WkNo, ret)
  dim StrSQL,Tmp1,Tmp2
  ret = true
  If Flag <> 4 Then
   If WFlag = 1 Then
'     StrSQL="SELECT Count(WkContrlNo) AS num from hITFullOutSelect "&_
'            "WHERE ContNo="& strNum
      StrSQL="SELECT Count(A.WkContrlNo) AS num from hITCommonInfo A LEFT JOIN hITFullOutSelect B "&_
             "ON A.WkContrlNo=B.WkContrlNo WHERE B.ContNo=" & strNum &" AND A.Process='R' "&_
             "AND A.WkType="& WFlag &"AND A.WorkCompleteDate Is Null"
'20030911 ADD This Item "AND A.WorkCompleteDate Is Null" 
     ObjRS.Open StrSQL, ObjConn
     if err <> 0 then
       DisConnDBH ObjConn, ObjRS	'DB切断
       jampErrerP "1","b000","00","IT共通テーブル","101","SQL：<BR>"&StrSQL
     end if
     Tmp1=ObjRS("num")
     ObjRS.close
   Else
     Tmp1=0
   End If
   If Tmp1 <> 0 Then
      StrSQL="SELECT WkNo from hITCommonInfo A LEFT JOIN hITFullOutSelect B "&_
             "ON A.WkContrlNo=B.WkContrlNo WHERE B.ContNo=" & strNum &" AND A.Process='R' "&_
             "AND A.WkType="& WFlag &"AND A.WorkCompleteDate Is Null"
'20030911 ADD This Item "AND A.WorkCompleteDate Is Null" 
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","IT共通テーブル","102","SQL：<BR>"&StrSQL
      end if
      If Not IsNull(ObjRS("WkNo")) Then
        WkNo=ObjRS("WkNo")
        ret =false
      End If
      ObjRS.close
    Else
      StrSQL="SELECT Count(WkContrlNo) AS num from hITCommonInfo "&_
             "WHERE ContNo="& strNum &" AND Process='R' AND WkType="& WFlag &" AND WorkCompleteDate Is Null"
'20030911 ADD This Item "AND WorkCompleteDate Is Null" 
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","IT共通テーブル","101","SQL：<BR>"&StrSQL
      end if
      Tmp2=ObjRS("num")
      ObjRS.close
      If Tmp2 <> 0 Then
        StrSQL="SELECT WkNo from hITCommonInfo WHERE ContNo="& strNum &" AND Process='R' "&_
                "AND WkType="& WFlag&"AND WorkCompleteDate Is Null"
'20030911 ADD This Item "AND WorkCompleteDate Is Null" 
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b000","00","IT共通テーブル","102","SQL：<BR>"&StrSQL
        end if
        WkNo=ObjRS("WkNo")
        ObjRS.close
        ret=false
      End If
    End If
  Else
    StrSQL="SELECT Count(WkContrlNo) AS num from hITCommonInfo WHERE BLNo="& strNum &" AND Process='R' "&_
           "AND WorkCompleteDate Is Null"
'20030911 ADD This Item "AND WorkCompleteDate Is Null" 
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      jampErrerP "1","b000","00","IT共通テーブル","101","SQL：<BR>"&StrSQL
    end if
    Tmp1=ObjRS("num")
    ObjRS.close
    If Tmp1 <> 0 Then
      StrSQL="SELECT WkNo from hITCommonInfo WHERE BLNo=" & strNum &" AND Process='R' "&_
             "AND WorkCompleteDate Is Null"
'20030911 ADD This Item "AND WorkCompleteDate Is Null" 
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b000","00","IT共通テーブル","102","SQL：<BR>"&StrSQL
      end if
      WkNo=ObjRS("WkNo")
      ObjRS.close
      ret=false
    End If
  End If
End Function
'
'個別ブッキングテーブルの重複チェック&ブッキング指示テーブルの重複チェック
'Input :DB接続Object、レコードセットObject、ブックNo、登録者コード、指示先コード
'      :登録済み指示先CSV、処理区分1、処理区分2、戻り値
'Output:ret(重複なし又は処理区分「D」=>true　重複有り=>false) 
'      :PFlag(処理区分)
Function checkSPBook(ObjConn,ObjRS,BookNo,USER,Code,strCodes,PFlag1,PFlag2,ret)
  StrSQL = "SELECT Count(BookNo) AS Num FROM SPBookInfo WHERE BookNo='"& BookNo &"' "
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b000","00","個別ブッキングテーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("Num") <> 0 Then
    ObjRS.close
    StrSQL = "SELECT Process FROM SPBookInfo WHERE BookNo='"& BookNo &"' "
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","個別ブッキングテーブル","102","SQL：<BR>"&StrSQL
    end if
    PFlag1=ObjRS("Process")
    If PFlag1="R" Then
      ObjRS.close
      StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
               "WHERE BookNo='"& BookNo &"' AND SenderCode='"& USER &"' "&_
               "AND Process='R'"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
      end if
      If ObjRS("Num") > 0 Then
        If ObjRS("Num")=1 Then
          PFlag1="1"
        Else
          PFlag1="3"
        End If
      Else
        PFlag1="2"
      End If
'      ret = false
    Else
      PFlag1="2"
'      ret = true
    End If
  Else
    PFlag1="0"
  End If
  ObjRS.close
'3th ADD ↓↓↓↓↓↓↓↓↓↓↓↓
  If Trim(Code)="" OR Code=Null Then
    Code=" "
  End If
  StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
           "WHERE BookNo='"& BookNo &"' AND SenderCode='"& USER &"' "&_
           "AND TruckerCode='" & Code & "'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("Num") <> 0 Then
    ObjRS.close
    StrSQL = "SELECT TruckerCode,Process FROM BookingAssign "&_
             "WHERE BookNo='"& BookNo &"' AND SenderCode='"& USER &"'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
    end if
    strCodes=""
    Do Until ObjRS.EOF
      If Trim(ObjRS("Process"))="R" Then
        If Trim(Code) = Trim(ObjRS("TruckerCode")) Then
          ret = false
        End If
        If Trim(ObjRS("TruckerCode"))="" Then
          strCodes=strCodes &"なし,"
        Else
          strCodes=strCodes & Trim(ObjRS("TruckerCode")) &","
        End If
      End If
      ObjRS.MoveNext
    Loop
    PFlag2="1"
  Else
    PFlag2="0"
  End If
  ObjRS.close
End Function
'
'作業管理番号取得
'Input :DB接続Object、レコードセットObject、ユーザ略称、作業管理番号
'OutPut:作業管理番号
Function getWkContrlNo(ObjConn,ObjRS,sUN,WkContrlNo)
    dim StrSQL
    dim tmpContrlNo		'ADD 20040107
    StrSQL = "SELECT CurrentVal FROM mAutoNumber WHERE TypeCode='12'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","作業管理番号テーブル","102","SQL：<BR>"&StrSQL
    end if
    WkContrlNo = ObjRS("CurrentVal")+1
    tmpContrlNo = WkContrlNo	'ADD 20040107
    ObjRS.close
    If Month(Now)*100000000 > WkContrlNo Then
      WkContrlNo = Month(Now)*100000000 +1
'20040107 Change    ElseIf Month(Now) = 1 Then
    ElseIf Month(Now) = 1 AND WkContrlNo/100000000>=12 Then
      WkContrlNo = 100000001
    '2007/02/01 Add-S M.Marquez  
    ElseIf Month(Now) <> 1 AND WkContrlNo/100000000>=12 Then
    '2007/02/01 Add-E M.Marquez  
		WkContrlNo = (Month(Now)*100000000) + 1
    End If
    dim ret
    ret= false
'ADD 作業管理番号重複時の処理 2003/08/04 ↓↓↓↓↓↓↓↓↓↓↓
    Do Until ret
     StrSQL = "SELECT COUNT(ITC.WkContrlNo) AS Expr1,COUNT(ITF.WkContrlNo) AS Expr2,COUNT(ITR.WkContrlNo) AS Expr3 "&_
               "FROM hITReference ITR "&_
               "FULL OUTER JOIN hITFullOutSelect ITF ON ITR.WkContrlNo = ITF.WkContrlNo "&_
               "FULL OUTER JOIN hITCommonInfo ITC ON ITF.WkContrlNo = ITC.WkContrlNo "&_
               "WHERE (ITC.WkContrlNo="& WkContrlNo &") OR (ITF.WkContrlNo="& WkContrlNo &") "&_
               "OR (ITR.WkContrlNo ="& WkContrlNo &")"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b000","00","作業管理番号テーブル","102","SQL：<BR>"&StrSQL
      end if
      If ObjRS("Expr1")+ObjRS("Expr2")+ObjRS("Expr3") = 0 Then
        ret = true
      Else
'20040107 Change        WkContrlNo = WkContrlNo +1
		'2007/02/01 Upd-S M.Marquez  
        'WkContrlNo = tmpContrlNo +1000
        WkContrlNo =GetMonthMaxWorkCtrlNo(ObjConn,100000000) + 1
		'2007/02/01 Upd-E M.Marquez
      End If
      ObjRS.close
    Loop
'ADD 作業管理番号重複時の処理 2003/08/04 ↑↑↑↑↑↑↑↑↑↑↑
    StrSQL = "UPDATE mAutoNumber SET CurrentVal = "& WkContrlNo &", UpdtTime='"& now() &"'," &_
             "UpdtPgCd='PREDEF01',UpdtTmnl='"& sUN &"' WHERE TypeCode='12'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn, "1","b000","00","作業管理番号テーブル","104","SQL：<BR>"&StrSQL
    end if
End Function
'2007/02/01 Add-S M.Marquez
'-------------------------------------------
'各月の作業管理番号の最大を取得
'	ObjConn = DB Connection
'   lStart = Start of Series 
'-------------------------------------------
Function GetMonthMaxWorkCtrlNo(ObjConnlocal,lStart)
	dim sql, rsMax
	dim lMax(2)
	on error resume next
	'DB接続
	ConnDBH ObjConnlocal, rsMax
	
	'Get the max value registered in hitReference table
	sql = "select isnull(Max(WKContrlNo),(Month(getdate())) * " & lStart & ") as WorkCtrl from hitReference "
	sql = sql & "where WKContrlNo between ((Month(getdate())) * " & lStart & ") and ((Month(getdate()) +1) * " & lStart & ")"
	rsMax.Open sql, ObjConn
	
	if not rsMax.eof then 
		lMax(0)=rsMax("WorkCtrl")
	else
		lMax(0)=(lStart * Month(Now)) + 1
	end if
	rsMax.Close
	
	'Get the max value registered in hitCommonInfo table
	sql = "select isnull(Max(WKContrlNo),(Month(getdate())) * " & lStart & ") as WorkCtrl from hitCommonInfo "
	sql = sql & "where WKContrlNo between ((Month(getdate())) * " & lStart & ") and ((Month(getdate()) +1) * " & lStart & ")"
	rsMax.Open sql, ObjConn
	
	if not rsMax.eof then 
		lMax(1)=rsMax("WorkCtrl")
	else
		lMax(1)=(lStart * Month(Now)) + 1
	end if
	rsMax.Close

	'Get the max value registered in hitFullOutSelect table
	sql = "select isnull(Max(WKContrlNo),(Month(getdate())) * " & lStart & ") as WorkCtrl from hitFullOutSelect "
	sql = sql & "where WKContrlNo between ((Month(getdate())) * " & lStart & ") and ((Month(getdate()) +1) * " & lStart & ")"
	rsMax.Open sql, ObjConn
	
	if not rsMax.eof then 
		lMax(2)=rsMax("WorkCtrl")
	else
		lMax(2)=(lStart * Month(Now)) + 1
	end if
	rsMax.Close
	if lMax(0) > lMax(1) then
		if lMax(0) > lMax(2) then
			GetMonthMaxWorkCtrlNo =  lMax(0)
		else
			GetMonthMaxWorkCtrlNo =  lMax(2)
		end if
	else
		if lMax(1) > lMax(2) then
			GetMonthMaxWorkCtrlNo =  lMax(1)
		else
			GetMonthMaxWorkCtrlNo =  lMax(2)
		end if
	end if
End Function
'2007/02/01 Add-E M.Marquez
'ヘッド会社コード存在チェック
'Input :DB接続Object、レコードセットObject、ヘッド会社コード配列、戻り値
'OutPut:ret(存在する=>true　存在しない=>false)
Function checkHdCd(ObjConn,ObjRS,CMPcd,ret)
  dim StrSQL,check,i
  check=0
  For i=1 to 4
    If CMPcd(i) <> "" And CMPcd(i) <> "**" Then
      StrSQL="SELECT count(*) AS num FROM mUsers Where HeadCompanyCode='"& CMPcd(i) &"'"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerP "1","b000","00","IT業者共通マスタテーブル","101","SQL：<BR>"&StrSQL
      end if
      If ObjRS("num") = 0 Then
        check = check+1
      End If
      ObjRS.close
    End If
  Next
  If check=0 Then
    ret=true
  Else
    ret=false
  End If
End Function

'
'INIファイルからパラメータを読込む
'Input :Array(1)
'OutPut:Array(0)=ALARM、Array(1)=COMPLETED_DATA
Function getIni(param)
  dim ObjFSO,ObjTS,tmpStr
  Set ObjFSO = Server.CreateObject("Scripting.FileSystemObject")
  '--- ファイルを開く（読み取り専用） ---
  Set ObjTS = ObjFSO.OpenTextFile(Server.Mappath("./INI/PREDEF.INI"),1,false)
  '--- ファイルデータの読込み ---
  Do Until ObjTS.AtEndofStream
    tmpStr = Split(ObjTS.ReadLine, "=", 3, 1)
    Select Case tmpStr(0)
      Case "ALARM"
        param(0) = tmpStr(1)
      Case "COMPLETED_DATA"
        param(1) = tmpStr(1)
      Case "INPUT_DATA"
        param(2) = tmpStr(1)
    End Select
  Loop
  ObjTS.Close
  Set ObjTS = Nothing
  Set ObjFSO = Nothing
End Function

'ヘッド会社コード表示制御
'Input :ヘッド会社コード配列、アップフラグ
'
Function chengeCompCd(CMPcd,UpFlag)
  dim hCd,check,User,i
  hCd    = Session.Contents("COMPcd")
  User   = Session.Contents("userid")
  check=Array(0,0,0,0,0,0)
    If CMPcd(0) = UCase(User) Then
      UpFlag = 1
      check(0)=1
      check(1)=1
    End If
    For i=1 to 4
      If CMPcd(i) = hCd Then
        UpFlag = i+1
        check(i-1)=1
        check(i)=1
        check(i+1)=1
      End If
    Next
    For i=1 to 4
      If check(i) = 0 Then
        CMPcd(i)="**"
      End If
    Next
End Function
'ヘッド会社コード入力制御
'Input :ヘッド会社コード配列、アップフラグ
'
Function saveCompCd(CMPcd,UpFlag)
  dim hCd,User,i
  hCd    = Session.Contents("COMPcd")
  User   = UCase(Session.Contents("userid"))
  Select Case UpFlag
    Case "5"
'CW-0037     If CMPcd(4) = hCd Then
      CMPcd(1) = "'"& CMPcd(1) &"' readOnly tabindex=-1"
      CMPcd(2) = "'"& CMPcd(2) &"' readOnly tabindex=-1"
      CMPcd(3) = "'"& CMPcd(3) &"' readOnly tabindex=-1"
      CMPcd(4) = "'"& CMPcd(4) &"' readOnly tabindex=-1"
    Case "4"
'CW-0037    ElseIf CMPcd(3) = hCd Then
      CMPcd(1) = "'"& CMPcd(1) &"' readOnly tabindex=-1"
      CMPcd(2) = "'"& CMPcd(2) &"' readOnly tabindex=-1"
      CMPcd(3) = "'"& CMPcd(3) &"' readOnly tabindex=-1"
'CW-0017      CMPcd(4) = "'"& CMPcd(4) &"'"
      CMPcd(4) = "'"& CMPcd(4) &"' onBlur='checkIDF(0)'"
    Case "3"
'CW-0037    ElseIf CMPcd(2) = hCd Then
      CMPcd(1) = "'"& CMPcd(1) &"' readOnly tabindex=-1"
      CMPcd(2) = "'"& CMPcd(2) &"' readOnly tabindex=-1"
'CW-0017      CMPcd(3) = "'"& CMPcd(3) &"'"
      CMPcd(3) = "'"& CMPcd(3) &"' onBlur='checkIDF(0)'"
      CMPcd(4) = "'"& CMPcd(4) &"' readOnly tabindex=-1"
    Case "2"
'CW-0037    ElseIf CMPcd(1) = hCd Then
      CMPcd(1) = "'"& CMPcd(1) &"' readOnly"
'CW-0017      CMPcd(2) = "'"& CMPcd(2) &"'"
      CMPcd(2) = "'"& CMPcd(2) &"' onBlur='checkIDF(0)'"
      CMPcd(3) = "'"& CMPcd(3) &"' readOnly tabindex=-1"
      CMPcd(4) = "'"& CMPcd(4) &"' readOnly tabindex=-1"
    Case "1"
'CW-0037    Else
'CW-0017      CMPcd(1) = "'"& CMPcd(1) &"'"
      CMPcd(1) = "'"& CMPcd(1) &"' onBlur='checkIDF(0)'"
      CMPcd(2) = "'"& CMPcd(2) &"' readOnly tabindex=-1"
      CMPcd(3) = "'"& CMPcd(3) &"' readOnly tabindex=-1"
      CMPcd(4) = "'"& CMPcd(4) &"' readOnly tabindex=-1"
    End Select
'CW-0037    End If
End Function

'データ受渡しメソッド		'CW-055 ADD
'Input :データ数、データテーブル
'
Function Set_Data(Num,DtTbl)
  dim i,j
  Response.Write "  <INPUT type=hidden name='DataNum' value='" & Num & "'>" & vbCrLf
  For i = 0 to Num
    Response.Write "  <INPUT type=hidden name='Datatbl" & i & "' value='"
    For j=0 To UBound(DtTbl(i))-1
      Response.Write DtTbl(i)(j) & ","
    Next
    Response.Write DtTbl(i)(j) & "'>" & vbCrLf
  Next
End Function

'データ受取りメソッド		'CW-055 ADD
'Input :データ数、データテーブル
'
Function Get_Data(Num,DtTbl)
  dim param
  Num=Request("DataNum")
  ReDim DtTbl(Num)
  For Each param In Request.Form
    If Left(param, 7) = "Datatbl" Then
      DtTbl(Mid(param,8))=Split(Request.Form(param), ",", -1, 1)
    End If
  Next
End Function

'作業番号採番		'3th ADD
'Input :DB接続Object、レコードセットObject、ユーザID、作業番号
'OutPut:作業番号
Function getWkNo(ObjConn,ObjRS,USER,WkContrlNo)
  dim StrSQL
  StrSQL = "SELECT Count(WkNo) AS Num FROM hITWkNo where Status='3'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    ObjRS.Close
    Set ObjRS = Nothing
    jampErrerPDB ObjConn, "1","b001","00","作業番号管理テーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("Num") <> 0 Then
    ObjRS.close
    '2007/12/13 Upd-S Y.Nakano
    'StrSQL = "SELECT WkNo FROM hITWkNo where Status='3'"
    StrSQL = "SELECT TOP 1 WkNo FROM hITWkNo WITH(READPAST) where Status='3' "
    '2007/12/13 Upd-E Y.Nakano
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      ObjRS.Close
      Set ObjRS = Nothing
      jampErrerPDB ObjConn, "1","b001","00","作業番号管理テーブル","102","SQL：<BR>"&StrSQL
    end if
    SakuNo=ObjRS("WkNo")
    ObjRS.close
    StrSQL = "UPDATE hITWkNo SET Status='2' WHERE WkNo ='"& SakuNo &"'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn, "1","b001","00","作業番号管理テーブル","104","SQL：<BR>"&StrSQL
    end if
  Else
    ObjRS.close
    StrSQL = "SELECT CurrentVal FROM mAutoNumber WHERE TypeCode='11'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      ObjRS.Close
      Set ObjRS = Nothing
      jampErrerPDB ObjConn, "1","b001","00","番号管理マスタ","102","SQL：<BR>"&StrSQL
    end if
    SakuNo = ObjRS("CurrentVal")+1
    ObjRS.close
    StrSQL = "UPDATE mAutoNumber SET CurrentVal = "& SakuNo &", UpdtTime='"& now() &"',"&_
             "UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"' WHERE TypeCode='11'"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn, "1","b001","00","作業番号管理テーブル","104","SQL：<BR>"&StrSQL
    end if
    SakuNo = Right("0000" & Hex(SakuNo),5)
    StrSQL = "Insert Into hITWkNo (WkNo,UpdtTime,UpdtPgCd,UpdtTmnl,Status) values ('" &_
             SakuNo &"','"& Now() &"','PREDEF01','"& USER &"','2')"
    ObjConn.Execute(StrSQL)
    if err <> 0 then
      Set ObjRS = Nothing
      jampErrerPDB ObjConn, "1","b001","00","作業番号管理テーブル","103","SQL：<BR>"&StrSQL
    end if
  End If
End Function
'
'ブッキング指示テーブルの重複チェック		'3th ADD
'Input :DB接続Object、レコードセットObject、ブックNo、登録者コード
'      :戻り値
'Output:ret(重複なし又は処理区分「D」=>true　重複有り=>false) 
Function BookAs(ObjConn,ObjRS,BookNo,USER,ret)
 '2016/08/25 H.Yoshikawa Upd Start
  'StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
  '         "WHERE BookNo='"& BookNo &"' AND Process='R'"
  StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
           "WHERE BookNo='"& BookNo &"' AND Process='R' AND SenderCode = '" & USER & "' "
  ObjRS.Open StrSQL, ObjConn
 '2016/08/25 H.Yoshikawa Upd End
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("Num")>0 Then
 '2016/08/25 H.Yoshikawa Upd Start
    ObjRS.close
    'StrSQL = "SELECT SenderCode,Process FROM BookingAssign "&_
    '         "WHERE BookNo='"& BookNo &"' AND Process='R'"
    'ObjRS.Open StrSQL, ObjConn
    'if err <> 0 then
    '  DisConnDBH ObjConn, ObjRS	'DB切断
    '    jampErrerP "1","b000","00","ブッキング指示テーブル","102","SQL：<BR>"&StrSQL
    'end if
    'If Trim(ObjRS("SenderCode")) <> USER Then
      ret=false
    'End If
 '2016/08/25 H.Yoshikawa Upd End
  End If
  ObjRS.close
End Function

'ブッキングの重複チェック'2006/03/06 h.matsuda add
'Input :ブックNo、
'      :戻り値 件数
Function BookCountSL(BookNo)
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS
  StrSQL = "select count(BOK.BookNo) as Num "&_
		   "from(select distinct BookNo,shipline from Booking) as BOK "&_
		   "where BOK.BookNo='"& BookNo & "' "

  ObjRS.Open StrSQL, ObjConn

  BookCountSL=Trim(ObjRS("Num"))

  ObjRS.Close

End Function

'ブッキング指示テーブルの重複チェック		'2006/03/06 h.matsuda add
'Input :DB接続Object、レコードセットObject、ブックNo、登録者コード
'      :戻り値
'Output:ret(重複なし又は処理区分「D」=>true　重複有り=>false) 
Function BookAs2(ObjConn,ObjRS,BookNo,USER,ret,ShipLine)
 '2016/08/25 H.Yoshikawa Upd Start
  'StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
  '         "WHERE BookNo='"& BookNo &"' AND Process='R'"
  StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
           "WHERE BookNo='"& BookNo &"' AND Process='R' "&_
           "  AND SenderCode = '" & USER & "' AND ShipLine = '" & ShipLine & "' "
 '2016/08/25 H.Yoshikawa Upd End
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("Num")>0 Then
   '2016/08/25 H.Yoshikawa Upd Start
    ObjRS.close
    'StrSQL = "SELECT SenderCode,Process,ShipLine FROM BookingAssign "&_
    '         "WHERE BookNo='"& BookNo &"' AND Process='R'"
    'ObjRS.Open StrSQL, ObjConn
    'if err <> 0 then
    '  DisConnDBH ObjConn, ObjRS	'DB切断
    '    jampErrerP "1","b000","00","ブッキング指示テーブル","102","SQL：<BR>"&StrSQL
    'end if
	
	'do until ObjRS.EOF
	'	If Trim(ObjRS("SenderCode")) = USER and _
	'			(Trim(ObjRS("ShipLine")) =ShipLine or Trim(ObjRS("ShipLine")) ="") Then'OK
	'	  exit do
	'	ElseIf Trim(ObjRS("SenderCode")) <> USER and _
	'			(Trim(ObjRS("ShipLine")) =ShipLine or Trim(ObjRS("ShipLine")) ="") Then'NG
    '     ret=false
	'	  exit do
	'	ElseIf Trim(ObjRS("SenderCode")) = USER and _
	'			Trim(ObjRS("ShipLine")) <>ShipLine and Trim(ObjRS("ShipLine")) <>"" Then'NG
    '      ret=false
	'	  exit do
	'	End If
	'	ObjRS.movenext
	'loop
	ret=false
  End If
  ObjRS.close
End Function

'個別ブッキングテーブルの重複チェック&ブッキング指示テーブルの重複チェック
'Input :DB接続Object、レコードセットObject、ブックNo、登録者コード、指示先コード
'      :登録済み指示先CSV、処理区分1、処理区分2、戻り値
'Output:ret(重複なし又は処理区分「D」=>true　重複有り=>false) 
'      :PFlag(処理区分)
'2006/03/06 add h.matsuda
Function checkSPBook2(ObjConn,ObjRS,BookNo,USER,Code,strCodes,PFlag1,PFlag2,ret)
  StrSQL = "SELECT Count(BookNo) AS Num FROM SPBookInfo WHERE BookNo='"& BookNo &"' "
  StrSQL= StrSQL & " and SenderCode = '" & USER & "'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b000","00","個別ブッキングテーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("Num") <> 0 Then
    ObjRS.close
    StrSQL = "SELECT Process FROM SPBookInfo WHERE BookNo='"& BookNo &"' "
    StrSQL= StrSQL & " and SenderCode = '" & USER & "'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","個別ブッキングテーブル","102","SQL：<BR>"&StrSQL
    end if
    PFlag1=ObjRS("Process")
    If PFlag1="R" Then
      ObjRS.close
      StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
               "WHERE BookNo='"& BookNo &"' AND SenderCode='"& USER &"' "&_
               "AND Process='R'"
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
      end if
      If ObjRS("Num") > 0 Then
        '2016/08/31 H.Yoshikawa Upd Start
        'If ObjRS("Num")=1 Then
        '  PFlag1="1"
        'Else
        '  PFlag1="3"
        'End If
        PFlag1="1"
        '2016/08/31 H.Yoshikawa Upd Start
      Else
        PFlag1="2"
      End If
    Else
      PFlag1="2"
    End If
  Else
    PFlag1="0"
  End If
  ObjRS.close
  If Trim(Code)="" OR Code=Null Then
    Code=" "
  End If
  StrSQL = "SELECT Count(BookNo) AS Num FROM BookingAssign "&_
           "WHERE BookNo='"& BookNo &"' AND SenderCode='"& USER &"' "&_
           "AND TruckerCode='" & Code & "'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
  end if
  If ObjRS("Num") <> 0 Then
    ObjRS.close
    StrSQL = "SELECT TruckerCode,Process FROM BookingAssign "&_
             "WHERE BookNo='"& BookNo &"' AND SenderCode='"& USER &"'"
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b000","00","ブッキング指示テーブル","101","SQL：<BR>"&StrSQL
    end if
    strCodes=""
    Do Until ObjRS.EOF
      If Trim(ObjRS("Process"))="R" Then
        If Trim(Code) = Trim(ObjRS("TruckerCode")) Then
          ret = false
        End If
        If Trim(ObjRS("TruckerCode"))="" Then
          strCodes=strCodes &"なし,"
        Else
          strCodes=strCodes & Trim(ObjRS("TruckerCode")) &","
        End If
      End If
      ObjRS.MoveNext
    Loop
    PFlag2="1"
  Else
    PFlag2="0"
  End If
  ObjRS.close
End Function
'2010/04/23 Add-S Tanaka
'---------------------------------------------------------------
'【処理内容】ソートでNullが入った項目を最後にするためのSQL文作成
'【パラメータ】
'              strSortItemName         ソート対象項目
'              intSortType             並び順(0:ASC 1:DESC)
'【戻り値】    変換後のSQL文
'---------------------------------------------------------------
Function ChangeSortItem(strSortItemName,intSortType)
    Dim strWkSqL

    '変数初期化
     strWkSqL=""

    '並び順がASCの場合
    IF intSortType=0 Then
        strWkSqL="           CASE " 
        strWkSqL=strWkSqL & " WHEN " & strSortItemName & " IS NULL THEN '9' "
        strWkSqL=strWkSqL & " WHEN " & strSortItemName & "='' THEN '9' "
        strWkSqL=strWkSqL & " ELSE  '1' "
        strWkSqL=strWkSqL & " END + " & strSortItemName
    ELSE
        strWkSqL="           CASE " 
        strWkSqL=strWkSqL & " WHEN " & strSortItemName & " IS NULL THEN '1' "
        strWkSqL=strWkSqL & " WHEN " & strSortItemName & "='' THEN '1' "
        strWkSqL=strWkSqL & " ELSE  '9' "
        strWkSqL=strWkSqL & " END + " & strSortItemName
    End IF

    ChangeSortItem=strWkSqL
End Function
'2010/04/23 Add-E Tanaka
Function ChangeSortItem1(strSortItemName)
    Dim strWkSqL

    '変数初期化
    strWkSqL=""

    strWkSqL="           CASE " 
    strWkSqL=strWkSqL & " WHEN LTRIM(RTRIM((ISNULL(" & strSortItemName & ",''))))='' THEN 1 "
    strWkSqL=strWkSqL & " ELSE  0 "
    strWkSqL=strWkSqL & " END " & strSortItemName & "Ord "

    ChangeSortItem1=strWkSqL
End Function

function gfPutPage(rec,page,pagecount,link)
	dim pg, i, j
	dim FirstPage, LastPage	
	dim PageIndex
	dim PageWkNo	
	PageIndex=0
	PageWkNo=0

	if rec > 0 then	
		'--- カレントページ（補正）
		if pagecount<page then
			page=pagecount
		end if
		
		'ページIndexを設定
		PageIndex=Fix(page/gcPage)
		if page mod gcPage=0 then
			PageIndex=PageIndex-1
		End If
		PageWkNo=((gcPage*PageIndex)+1)
		
		'4/10
		'PageIndex = 0
		
		'先頭ページが0より小さい場合は1を設定
		if PageWkNo<=0 Then
			PageWkNo=0
		End If		

		'--- 総件数、総ページ数 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & FirstPage & """>最初へ</a>"
			response.write "| &nbsp;"											
			'if PageWkNo<>0 Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & page - 1 & """>前へ</a>"
			'Else
			'	response.write "<font style='color:#FFFFFF;'>前へ</font>"
			'End If
		else
			response.write "<font style='color:#FFFFFF;'>最初へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>前へ</font>"
		end if        

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
					response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & """>&nbsp;" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If		

		if page<=pagecount-1 then
			PageWkNo=page+1					
			'If PageWkNo<LastPage Then
				response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & PageWkNo & """>次へ</a>"'
			'Else
			'	response.write "<font style='color:#FFFFFF;'>次へ</font>"
			'End If
			response.write "| &nbsp;"
			response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & link & "=" & LastPage & """>最後へ</a>"'        
		else
			response.write "<font style='color:#FFFFFF;'>次へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>最後へ</font>"
		end if
	end if
end function
function gfPutPageSort(rec,page,pagecount,link,sortflag)
	dim pg, i, j
	dim FirstPage, LastPage	
	dim PageIndex
	dim PageWkNo	
	PageIndex=0
	PageWkNo=0

	if rec > 0 then	
		'--- カレントページ（補正）
		if pagecount<page then
			page=pagecount
		end if
		
		'ページIndexを設定
		PageIndex=Fix(page/gcPage)
		if page mod gcPage=0 then
			PageIndex=PageIndex-1
		End If
		PageWkNo=((gcPage*PageIndex)+1)
		
		'4/10
		'PageIndex = 0
		
		'先頭ページが0より小さい場合は1を設定
		if PageWkNo<=0 Then
			PageWkNo=0
		End If		

		'--- 総件数、総ページ数 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			'response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?SortFlag=" & sortflag & "&" & link & "=" & FirstPage & """>最初へ</a>"
			response.write "<a href='#' onclick='document.forms[0].pagenum.value=" & FirstPage & ";document.forms[0].submit();return false;'>最初へ</a>"
			response.write "| &nbsp;"											
			response.write "<a href='#' onclick='document.forms[0].pagenum.value=" & page - 1 & ";document.forms[0].submit();return false;'>前へ</a>"
			
		else
			response.write "<font style='color:#FFFFFF;'>最初へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>前へ</font>"
		end if        

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
					'response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?SortFlag=" & sortflag & "&" & link & "=" & PageWkNo & """>&nbsp;" & PageWkNo & "</a>"
					response.write "&nbsp;<a href='#' onclick='document.forms[0].pagenum.value=" & PageWkNo & ";document.forms[0].submit();return false;'>" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If		

		if page<=pagecount-1 then
			PageWkNo=page+1			
			response.write "<a href='#' onclick='document.forms[0].pagenum.value=" & PageWkNo & ";document.forms[0].submit();return false;'>次へ</a>"
			response.write "| &nbsp;"
			'response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?SortFlag=" & sortflag & "&" & link & "=" & LastPage & """>最後へ</a>"'        
			response.write "<a href='#' onclick='document.forms[0].pagenum.value=" & LastPage & ";document.forms[0].submit();return false;'>最後へ</a>"
		else
			response.write "<font style='color:#FFFFFF;'>次へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>最後へ</font>"
		end if
	end if
end function
'2015-01-23 Y.TAKAKUWA Add-S
function gfPutPageSort2(rec,page,pagecount,link,sortflag)
	dim pg, i, j
	dim FirstPage, LastPage	
	dim PageIndex
	dim PageWkNo	
	PageIndex=0
	PageWkNo=0

	if rec > 0 then	
		'--- カレントページ（補正）
		if pagecount<page then
			page=pagecount
		end if
		
		'ページIndexを設定
		PageIndex=Fix(page/gcPage)
		if page mod gcPage=0 then
			PageIndex=PageIndex-1
		End If
		PageWkNo=((gcPage*PageIndex)+1)
		
		'4/10
		'PageIndex = 0
		
		'先頭ページが0より小さい場合は1を設定
		if PageWkNo<=0 Then
			PageWkNo=0
		End If		

		'--- 総件数、総ページ数 
		LastPage=pagecount		
		FirstPage=1
			
		if page>1 then
			'response.write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?SortFlag=" & sortflag & "&" & link & "=" & FirstPage & """>最初へ</a>"
			response.write "<a href='#' onclick='showPage(" & FirstPage & ")'>最初へ</a>"
			response.write "| &nbsp;"											
			response.write "<a href='#' onclick='showPage(" & page - 1 & ")'>前へ</a>"
		else
			response.write "<font style='color:#FFFFFF;'>最初へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>前へ</font>"
		end if        

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
   				    response.write "&nbsp;<a href='#' onclick='showPage(" & PageWkNo & ")'>" & PageWkNo & "</a>"
				End If
			Next
			response.write "| &nbsp;"
		End If		

		if page<=pagecount-1 then
			PageWkNo=page+1				
			response.write "<a href='#' onclick='showPage(" & PageWkNo & ")'>次へ</a>"
			response.write "| &nbsp;"     
			response.write "<a href='#' onclick='showPage(" &  LastPage & ")'>最後へ</a>"
		else
			response.write "<font style='color:#FFFFFF;'>次へ</font>"
			response.write "| &nbsp;"
			response.write "<font style='color:#FFFFFF;'>最後へ</font>"
		end if
	end if
end function
'2015-01-23 Y.TAKAKUWA Add-E
'-----------------------------
'   日付編集（yy/mm/dd）
'-----------------------------
function gfformatDateS(d)
    dim w
    if isnull(d) then
        exit function
    end if
    gfformatDateS=""
    w=""
    if isdate(d) then
        w=year(d)*10000 + month(d)*100 + day(d)
    elseif len(d)=10 then
        w=""
        gfformatDate=d
    elseif len(d)=8 then
        w=gfTrim(d)
    elseif len(d)=14 then
        w=left(d, 8)
    end if

    if w<>"" then
        gfformatDateS= mid(w,3,2) & "/" & mid(w,5,2) & "/" & mid(w,7,2)
    end if
end function

'Y.TAKAKUWA Add-S 2013-03-18
'Menuバーの描画
Sub DispMenuBar
    Response.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
    Response.Write "  <tr>"
    Response.Write "    <td valign='bottom' align='right'><a href='http://www.hits-h.com/index.asp'><img src='../gif/b-home.gif' border='0' width='270' height='23' usemap='#map'></a></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr>"
    Response.Write "   <td bgcolor='000099' height='10'><img src='../gif/1.gif'></td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

'戻り先の登録２（フレーム版）
Sub DispMenuBarBack2(strBackURL)
'    ' 戻り先を登録
'    Dim strBackURL
'    strBackURL = Trim(Session.Contents("returnurl"))
'    If strBackURL="" Then
'        strBackURL = "index.asp"
'    End If

    Response.Write "<map name='map'>"
    Response.Write "<area shape='poly' coords='20,0,152,0,134,22,0,22' href='" & strBackURL & "' target='_top'>"
    Response.Write "<area shape='poly' coords='154,0,136,22,284,22,284,0' href='http://www.hits-h.com/index.asp' target='_top'>"
    Response.Write "</map>"

'    Session.Contents("returnurl") = Request.ServerVariables("SCRIPT_NAME")
End Sub
'Y.TAKAKUWA Add-E 2013-03-18

%>