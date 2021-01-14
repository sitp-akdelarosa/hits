<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										   _/
'_/	FileName	:dmi320.asp									   _/
'_/	Function	:事前実搬入入力更新							   _/
'_/	Date		:2003/05/29									   _/
'_/	Code By		:SEIKO Electric.Co 大重						   _/
'_/	Modify		:C-002	2003/08/08	備考欄追加				   _/
'_/	Modify		:3th	2003/01/31	3次変更					   _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
  WriteLogH "b402", "実搬入事前情報入力","14",""

'サーバ日付の取得
  dim DayTime, YY,Yotei
  getDayTime DayTime

'ユーザデータ所得
  dim USER,sUN, Utype
  USER   = UCase(Session.Contents("userid"))
  sUN    = Session.Contents("sUN")
  Utype  = Session.Contents("UType")
'データ取得
  dim UpFlag,CONnum,SakuNo,BookNo
  dim CMPcd,HedId,HTo,Hmon,Hday,TuSk
  dim FullName,Mord,i
  dim SendUser
  Mord   = Request("Mord")
  UpFlag = Request("UpFlag")
  SakuNo = Request("SakuNo")
  CONnum = Request("CONnum")
  BookNo = Request("BookNo")
  CMPcd  = Array(Request("CMPcd0"),Request("CMPcd1"),Request("CMPcd2"),Request("CMPcd3"),Request("CMPcd4"))
  HedId   = Request("HedId")
  Hmon    = Right("00" & Request("Hmon") ,2)
  Hday    = Right("00" & Request("Hday") ,2)
  '作業予定日の年度を決定
  If DayTime(1) > Hmon Then	'来年
    YY = DayTime(0) +1
  ElseIf DayTime(1) = Hmon AND DayTime(2) > Hday Then	'CW-043
    YY = DayTime(0) +1					'CW-043
  Else
    YY = DayTime(0)
  End If
  If Hmon = "00" Or Hday = "00" Then
    Yotei= "Null"
  Else
    Yotei=  "'"& YY &"/"& Hmon &"/"& Hday &"'"
  End If
  If HedId = "" Then
    HedId   = "Null"
  Else
    HedId = "'"& HedId &"'"
  End If
'通関
  TuSk=Request("TuSk")
  If TuSk="済" Then
    TuSk="Y"
  Else
    TuSk="N"
  End If
  FullName= ""
'3th ADD ↓↓↓↓↓↓↓
  dim OH,OWL,OWR,OLF,OLA
  If Request("OH") <>"" Then OH =Request("OH")  Else OH ="Null"
  If Request("OWL")<>"" Then OWL=Request("OWL") Else OWL="Null"
  If Request("OWR")<>"" Then OWR=Request("OWR") Else OWR="Null"
  If Request("OLF")<>"" Then OLF=Request("OLF") Else OLF="Null"
  If Request("OLA")<>"" Then OLA=Request("OLA") Else OLA="Null"
'3th ADD ↑↑↑↑↑↑↑

 dim TruckerSubName
 TruckerSubName = Request("TruckerSubName")
 
'エラートラップ開始
'  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL,tmpStr 
  ConnDBH ObjConn, ObjRS

  dim ret,ErrerM
  ret = true
'3th ADD START  ↓↓↓↓↓↓↓
  If Mord = 0 Then	'新規登録
    dim WkContrlNo,UpdateFlag,RFlag
    RFlag=0
    '重複登録チェック
    StrSQL = "SELECT Count(ITC.WkContrlNo) AS Num "&_
             "FROM hITCommonInfo AS ITC LEFT JOIN CYVanInfo AS CYV ON (ITC.WkNo = CYV.WkNo) AND (ITC.ContNo = CYV.ContNo) "&_
             "WHERE ITC.ContNo='" & CONnum &"' AND ITC.Process='R' AND ITC.WkType='3' AND CYV.BookNo='"& BookNo &"' "
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b401","03","実搬入：重複チェック","101","SQL:<BR>"&StrSQL
    end if
    If Trim(ObjRS("Num")) <> "0" Then
      ret=false
      ErrerM="操作中に指定したブッキングNo、コンテナ番号が登録されました。<BR>このため登録処理はキャンセルされます</P>"
    End If
    SendUser = CMPcd(1)
    ObjRS.Close
    If ret Then
      'CYVaninfoテーブルに過去データが残っているかチェック
      StrSQL = "SELECT Count(CYV.ContNo) AS Num "&_
               "FROM CYVanInfo AS CYV "&_
               "WHERE CYV.ContNo='" & CONnum &"' AND CYV.SenderCode='" & USER &"' AND CYV.BookNo='"& BookNo &"' "
      ObjRS.Open StrSQL, ObjConn
      if err <> 0 then
        DisConnDBH ObjConn, ObjRS	'DB切断
        jampErrerP "1","b401","03","実搬入：CYVaninfoテーブルチェック","101","SQL:<BR>"&StrSQL
      end if
      If Trim(ObjRS("Num")) <> "0" Then
        UpdateFlag = true
      Else
        UpdateFlag = false
      End If
      ObjRS.Close
      '元請陸運業者名取得
      If CMPcd(1) <> "" Then
        StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & CMPcd(1) &"'"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          DisConnDBH ObjConn, ObjRS	'DB切断
          jampErrerPDB ObjConn,"1","b401","03","実搬入：データ登録","102","SQL:<BR>"&StrSQL
        end if
        FullName = "'" & ObjRS("FullName") & "' "
        ObjRS.close
      Else
        FullName = "Null"
      End If
      'データ整形
      For i=1 To 4
        If CMPcd(i) = "" Then
          CMPcd(i) = "Null"
        Else
          If CMPcd(i) = USER Then
            RFlag=1
          End If
          CMPcd(i) = "'" & CMPcd(i) & "'"
        End If
      Next
      '作業管理番号の採番
      getWkContrlNo ObjConn, ObjRS, sUN, WkContrlNo
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b401","03","実搬入：データ登録","103","作業管理番号取得に失敗<BR>"&StrSQL
      end if
      '作業番号の採番
      getWkNo ObjConn, ObjRS, USER, SakuNo
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b401","03","実搬入：データ登録","103","作業番号取得に失敗<BR>"&StrSQL
      end if
      'IT共通テーブルへの登録
        StrSQL = "Insert Into hITCommonInfo (WkContrlNo,UpdtTime,UpdtPgCd,UpdtTmnl,Status, " &_
                 "Process,WkType,InPutDate,UpdtUserCode,WkNo,ContNo,RegisterType,RegisterName, " &_
                 "RegisterCode,TruckerSubCode1,HeadID,WorkDate,TruckerName,Comment1,Comment2,Comment3,TruckerSubName1) "&_
                 "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"', "&_
                 "'0','R','3','"& Now() &"','"& USER &"','"& SakuNo &"','"& CONnum &"', "&_
                 "'"& Utype &"','"& sUN &"','"& CMPcd(0) &"',"& CMPcd(1) &","& HedId &", "&_
                 Yotei &","& FullName &",'"& Request("Comment1") &"','"& Request("Comment2") &"', "&_
                 "'"& Request("Comment3") &"','" & TruckerSubName & "'"&  ") "
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b401","03","実搬入：データ登録","103","SQL:<BR>"&StrSQL
        end if
      '照会テーブル登録
      StrSQL = "Insert Into hITReference (WkContrlNo, UpdtTime, UpdtPgCd,UpdtTmnl," &_
               "TruckerFlag1,TruckerFlag2,TruckerFlag3,TruckerFlag4)" &_
               "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"'," &_
               "'"&RFlag&"','0','0','0')"
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b401","03","実搬入：データ登録","103","SQL:<BR>"&StrSQL
      end if
      If UpdateFlag Then
        'CYVaninfoテーブルの更新
        StrSQL = "UPDATE CYVanInfo SET ContSize='"&Request("CONsize")&"', ContType='"&Request("CONtype")&"', "&_
                 "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
                 "ContHeight='"&Request("CONhite")&"', Material='"&Request("CONsitu")&"', "&_
                 "ShipLine='"&Request("ThkSya")&"',VslName='"&Request("ShipN")&"',"&_
                 "TareWeight="&Request("CONtear")&", CustOK='"&Request("MrSk")&"', "&_
                 "SealNo='"&Request("SealNo")&"', ContWeight="&Request("GrosW")&", "&_
                 "ReceiveFrom='"&Request("HFrom")&"', CustClear='"&TuSk&"', "&_
                 "Voyage='"&Request("NextV")&"',DPort='"&Request("AgeP")&"',"&_
                 "OvHeight="&OH&", OvWidthL="&OWL&",OvWidthR="&OWR&", OvLengthF="&OLF&", "&_
                 "OvLengthA="&OLA&",DelivPlace='"&Request("NiwataP")&"',"&_
                 "Operator='"&Request("Operator")&"', WkNo='"& SakuNo &"' "&_
                 "WHERE BookNo='"& BookNo &"' AND SenderCode='" & USER &"' AND ContNo='"& CONnum &"'  "
        ObjConn.Execute(StrSQL)
        if err <> 0 then
           Set ObjRS = Nothing
           jampErrerPDB ObjConn,"1","b401","03","実搬入：データ登録","104","SQL:<BR>"&StrSQL
        end if
      Else
        'CYVaninfoテーブルへの登録
        StrSQL = "Insert Into  CYVanInfo (SenderCode,BookNo,ContNo,UpdtTime,UpdtPgCd,UpdtTmnl, "&_
                 "ContSize,ContType,ContHeight,ShipLine,VslName,Voyage,DPort,DelivPlace, "&_
                 "SealNo,ContWeight,CustClear,Material,TareWeight,CustOK,ReceiveFrom, "&_
                 "OvHeight,OvWidthL,OvWidthR,OvLengthF,OvLengthA,Operator,WkNo) "&_
                 "values ('" & USER &"','"& BookNo &"','"& CONnum &"','"& Now() &"','PREDEF01','"& USER &"', "&_
                 "'"&Request("CONsize")&"','"&Request("CONtype")&"','"&Request("CONhite")&"', "&_
                 "'"&Request("ThkSya")&"','"&Request("ShipN")&"','"&Request("NextV")&"', "&_
                 "'"&Request("AgeP")&"','"&Request("NiwataP")&"','"&Request("SealNo")&"',"&_
                 "'"&Request("GrosW")&"','"&TuSk&"','"&Request("CONsitu")&"',"&Request("CONtear")&", " &_
                 "'"&Request("MrSk")&"','"&Request("HFrom")&"', "&_
                 OH&", "&OWL&","&OWR&","&OLF&","&OLA&", "&_
                 "'"&Request("Operator")&"','"& SakuNo &"')"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
           Set ObjRS = Nothing
           jampErrerPDB ObjConn,"1","b401","03","実搬入：データ登録","104","SQL:<BR>"&StrSQL
        end if
      End If
    End If
  Else
'3th ADD END  ↑↑↑↑↑↑↑
'CW-006	ADD START ↓↓↓↓↓↓↓
   '完了・更新チェック
    If UpFlag <>5 Then
      StrSQL="SELECT ITC.WorkCompleteDate, ITR.TruckerFlag"& UpFlag &" AS Flag "&_
             "FROM hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
             "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"
    Else
      StrSQL="SELECT WorkCompleteDate FROM hITCommonInfo " &_
             "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"
    End If
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      ObjRS.Close
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b402","14","実搬入：データ登録","102","SQL:<BR>"&StrSQL
    end if
    If NOT IsNull(ObjRS("WorkCompleteDate")) Then 
      ret=false
      ErrerM="指定の作業は画面操作中に作業が完了したため、更新はキャンセルされました。"
    End If
'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
    If Len(Request("partFlg"))=1 Then
      ObjRS.close
      StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
               "UpdtTmnl='"& USER &"', Status='0',Process='R',UpdtUserCode='"& USER &"', "&_
               "WorkDate=" & Yotei &_
               " Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b402","14","実搬入：データ登録","104","SQL:<BR>"&StrSQL
      end if
      StrSQL = "UPDATE CYVanInfo SET "&_
               "UpdtTime='"& Now() &"',UpdtPgCd='PREDEF01',UpdtTmnl='"& USER &"',"&_
               "SealNo='"&Request("SealNo")&"', ContWeight="&Request("GrosW")&", "&_
               "CustClear='"&TuSk&"' "&_
               "WHERE BookNo='"& BookNo &"' AND ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' "
'               "TareWeight="&Request("CONtear")
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b402","14","実搬入：データ登録","104","SQL:<BR>"&StrSQL
      end if
    Else
'ADD 20050303 END
     'チェック
      If UpFlag <>5 Then
        If Trim(ObjRS("Flag"))=1 Then 
          ret=false
          ErrerM="指定の作業は画面操作中に指示先に受諾されたため、更新はキャンセルされました。"
        End If
      End If
      ObjRS.close
      If ret Then
'CW-006	End ADD ↑↑↑↑↑↑↑
      'データ更新
        If Mord <> 2 Then	'更新
          If UpFlag<>5 Then
            If CMPcd(UpFlag)="" Then
              tmpStr=", TruckerSubCode"& UpFlag &"=Null "
            Else
              tmpStr=", TruckerSubCode"& UpFlag &"='"& CMPcd(UpFlag) & "' "
              SendUser = CMPcd(UpFlag)
            End If
          Else
            tmpStr=" "
          End If

          tmpStr = tmpStr & ", TruckerSubName"& UpFlag &"='"& TruckerSubName & "' "

        '元請陸運業者名取得
          If UpFlag<2 Then
            If CMPcd(1) <> "" Then
              StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & CMPcd(1) &"'"
              ObjRS.Open StrSQL, ObjConn
              if err <> 0 then
                DisConnDBH ObjConn, ObjRS	'DB切断
                jampErrerPDB ObjConn,"1","b402","14","実搬入：データ登録","102","SQL:<BR>"&StrSQL
              end if
              FullName = ",TruckerName='" & ObjRS("FullName") & "' "
              ObjRS.close
            Else
              FullName = ",TruckerName=Null "
            End If
          End If

          StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"', Status='0', Process='R', " &_
                   "UpdtUserCode='"& USER &"', "&_
                   "HeadID=" & HedId &", WorkDate=" & Yotei & tmpstr & FullName &_
                   ", Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"',Comment3='"& Request("Comment3") &"' "&_
                   "Where WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"
'C-002 ADD This Line : ", Comment1='"& Request("Comment1") &"',Comment2='"& Request("Comment2") &"',Comment3='"& Request("Comment3") &"' "&_
          ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b402","14","実搬入：データ登録","104","SQL:<BR>"&StrSQL
          end if
          If UpFlag = 5 Then
            tmpStr = " "
          Else
            If UpFlag = 1 AND CMPcd(1) = UCase(Session.Contents("COMPcd")) Then 
              tmpStr = ", TruckerFlag1=1 "
            Else
              tmpStr = ", TruckerFlag"& UpFlag &"=0 "
            End If
          End If
          UpFlag = UpFlag-1
          If UpFlag = 0 Then
            tmpStr = tmpStr&" "
          Else
            tmpStr = tmpStr&", TruckerFlag"& UpFlag &"=1 "
          End If
       '参照フラグ更新
          StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                   "UpdtTmnl='"& USER &"'"&tmpStr&_
                   "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                   "WHERE WkNo='"& SakuNo &"' AND Process='R' AND WkType='3')"
          ObjConn.Execute(StrSQL)
            if err <> 0 then
              Set ObjRS = Nothing
              jampErrerPDB ObjConn,"1","b402","14","実搬入：データ登録","104","SQL:<BR>"&StrSQL
            end if
          StrSQL = "UPDATE CYVanInfo SET ContSize='"&Request("CONsize")&"', ContType='"&Request("CONtype")&"', "&_
                   "ContHeight='"&Request("CONhite")&"', Material='"&Request("CONsitu")&"', "&_
                   "TareWeight="&Request("CONtear")&", CustOK='"&Request("MrSk")&"', "&_
                   "SealNo='"&Request("SealNo")&"', ContWeight="&Request("GrosW")&", "&_
                   "ReceiveFrom='"&Request("HFrom")&"', CustClear='"&TuSk&"', "&_
                   "OvHeight="&OH&", OvWidthL="&OWL&", OvWidthR="&OWR&", OvLengthF="&OLF&", OvLengthA="&OLA&" "&_
                   "WHERE BookNo='"& BookNo &"' AND ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' "
          ObjConn.Execute(StrSQL)
          if err <> 0 then
             Set ObjRS = Nothing
             jampErrerPDB ObjConn,"1","b402","14","実搬入：データ登録","104","SQL:<BR>"&StrSQL
          end if
        Else		'保留
          'ヘッダID更新
            If UpFlag=5 Then
              tmpStr=""
            Else
              tmpStr=", TruckerSubCode"& UpFlag &"=Null"
            End If
            StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                     "UpdtTmnl='"& USER &"', Status='0', Process='R', " &_
                     "UpdtUserCode='"& USER &"'"& tmpStr &", HeadID=Null " &_
                     "Where ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' AND Process='R' AND WkType='3'"
            ObjConn.Execute(StrSQL)
            if err <> 0 then
              Set ObjRS = Nothing
             jampErrerPDB ObjConn,"1","b402","15","実搬入：保留","102","SQL:<BR>"&StrSQL
            end if
           '参照フラグ更新
            StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                     "UpdtTmnl='"& USER &"', TruckerFlag"& UpFlag-1 &"=2 "&_
                     "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                     "WHERE ContNo='"& CONnum &"' AND WkNo='"& SakuNo &"' AND Process='R' AND WkType='3')"
            ObjConn.Execute(StrSQL)
            if err <> 0 then
              Set ObjRS = Nothing
              jampErrerPDB ObjConn,"1","b402","15","実搬入：保留","102","SQL:<BR>"&StrSQL
            end if
          End If
      End If		'CW-006
    End If		'ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
  End If
  
'データ取得
	Dim Email1, Email2, Email3, Email4, Email5
	Dim UserName,ComInterval,rc

	'''通信間隔取得
	StrSQL = "SELECT ComInterval FROM mParam WHERE Seq = '1'"

	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
	'''DB切断
		DisConnDBH ObjConn, ObjRS
		jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","実搬入：メール送信","104","SQL:<BR>"&StrSQL
	end if

	ComInterval = ObjRS("ComInterval")
	ObjRS.Close

	if SendUser <> "" then
	''作業発生配信情報の取得
		StrSQL = "SELECT T.*, "
		StrSQL = StrSQL & "CASE WHEN U.NameAbrev IS NULL THEN U.FullName ELSE U.NameAbrev END AS USERNAME "
		StrSQL = StrSQL & "FROM mUsers U, "
		StrSQL = StrSQL & "(SELECT T.* FROM TargetOperation T, mUsers U WHERE T.UserCode = U.UserCode "
		StrSQL = StrSQL & "AND U.HeadCompanyCode ='" & SendUser & "') T "
		StrSQL = StrSQL & "WHERE U.UserCode = '" & USER & "'"
		
		ObjRS.Open StrSQL, ObjConn
	    if ObjRS.EOF <> True then
		if err <> 0 then
	'''DB切断
			DisConnDBH ObjConn, ObjRS
			jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","実搬入：メール送信","104","SQL:<BR>"&StrSQL
		end if

		Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
		Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4, mailFlag5
		Dim mailSubject, mailBody,WorkName
		Dim SendTime, UpdateSendTime
		Dim fp, fobj, tfile

<!-- 2009/03/10 R.Shibuta Add-S -->
	'''SMTPサーバ名の設定
'		svName   = "slitdns2.hits-h.com"
		svName   = "192.168.17.61"
		attachedFiles = ""
		mailFlag1 = 0
		mailFlag2 = 0
		mailFlag3 = 0
		mailFlag4 = 0
		mailFlag5 = 0
	'''メール送信元アドレスの設定
		mailFrom = "mrhits@hits-h.com"
		mailTo = ""
		rc = ""
		
		if Trim(ObjRS("Email1")) <> "" AND ObjRS("FlagRecResults1") = "1" then
			mailTo = mailTo & Trim(ObjRS("Email1"))
			mailFlag1 = 1
		else
			mailFlag1 = 0
		end if

		if Trim(ObjRS("Email2")) <> "" AND ObjRS("FlagRecResults2") = "1" then
			if mailFlag1 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email2"))
			else
				mailTo = mailTo & Trim(ObjRS("Email2"))
			end if
				mailFlag2 = 1
		else
			mailFlag2 = 0
		end if

		if Trim(ObjRS("Email3")) <> "" AND ObjRS("FlagRecResults3") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
			else
				mailTo = mailTo & Trim(ObjRS("Email3"))
			end if
			mailFlag3 = 1
		else
			mailFlag3 = 0
		end if

		if Trim(ObjRS("Email4")) <> "" AND ObjRS("FlagRecResults4") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
			else
				mailTo = mailTo & Trim(ObjRS("Email4"))
			end if
			mailFlag4 = 1
		else
			mailFlag4 = 0
		end if

		if Trim(ObjRS("Email5")) <> "" AND ObjRS("FlagRecResults5") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
			else
				mailTo = mailTo & Trim(ObjRS("Email5"))
			end if
			mailFlag5 = 1
		else
			mailFlag5 = 0
		end if

		Set ObjMail = Server.CreateObject("BASP21")

		mailSubject = "HiTS 作業依頼"
		mailBody = "実搬入作業" & "発生 (" & Trim(ObjRS("USERNAME")) & "様より)" & vbCrLf & vbCrLf
		mailBody = mailBody & "実搬入作業" & "が発生しました。" & vbCrLf
		mailBody = mailBody & "詳しくはHiTSの事前情報登録の画面をご参照下さい。"
			
		'メール送信時刻から現在の時刻が通信間隔以上の場合はメールを送信する。

		
		if Trim(mailTo) <> "" Then
			if ObjRS("RecResultsDate") < DateAdd("n",(ComInterval * -1), Now())  OR IsNull(ObjRS("RecResultsDate")) = True then
				rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
				sendTime=Now
			end if

			If rc = "" Then
				'''メール送信日付の更新を行う。
				StrSQL = "UPDATE TargetOperation SET UpdtTime='" & Now() & "', UpdtPgCd='dmi340',"
				StrSQL = StrSQL & " UpdtTmnl='" & USER & "',"&  "RecResultsDate='" & Now() & "'"
				StrSQL = StrSQL &"WHERE UserCode = '" & Trim(ObjRS("UserCode")) & "'"

				ObjConn.Execute(StrSQL)
				if err <> 0 then
					Set ObjRS = Nothing
					jumpErrorPDB ObjConn,"1","c104","14","実搬入：メール送信","104","SQL:<BR>"&StrSQL
				end if
			else
			WriteLogH "b10", "こっちです", "",""
				fp = Server.MapPath("./mailerror") & "\error.txt"
				set fobj = Server.CreateObject("Scripting.FileSystemObject")
					if rc<>"" then
						if fobj.FileExists(fp) = True then
							set tfile = fobj.OpenTextFile(fp,8)
						else
							set tfile = fobj.CreateTextFile(fp,True,False)
						end if
						tfile.WriteLine sendTime & " " & rc
						tfile.Close
						ErrerM = "メール送信に失敗しました。<BR>"
						ret = 1
					end if
			end if
		else

		end if
	    end if
<!-- 2009/03/10 R.Shibuta Add-E -->
	end if
	
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<% If Mord =0 AND ret Then %>
<!-------------事前実搬入作業番号発行--------------------------->
<TITLE>作業番号発行</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
   try{
     window.resizeTo(200,300);
     window.opener.parent.List.location.href="./dmo310F.asp"
   }catch(e){
   }
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
  <P align=center><B>作業番号発行</B></P>
  <BR>
  <P>作業番号は「<%=SakuNo%>」です。</P>
  <BR><P><BR><P><BR><P>
  <P align=center><INPUT type=button value="閉じる" onClick="window.close()"></P>
<% ELSE %>
<!-------------事前実搬入入力更新--------------------------->
<TITLE>事前実搬入入力更新</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR><TD align=center>
  <% If ret Then%>
   更新しました。<BR>画面は自動的に閉じられます。
    <SCRIPT language=JavaScript>
     try{
       window.opener.parent.DList.location.href="./dmo310L.asp"
       window.opener.parent.Top.location.href="./dmo310T.asp"
     }catch(e){
     }
     window.close();
    </SCRIPT>
  <% Else %>
   <DIV class=alert><%=ErrerM%></DIV><BR>
   <INPUT type=button value="閉じる" onClick="window.close()">
  <% End If%>
    </TD></TR>
</TABLE>
<% End If %>
<!-------------画面終わり--------------------------->
</BODY></HTML>