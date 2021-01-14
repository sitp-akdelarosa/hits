<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi140.asp				_/
'_/	Function	:事前空搬入登録・更新			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/07/29	備考欄追加	_/
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

'サーバ日付の取得
  dim DayTime, YY,Yotei
  getDayTime DayTime

'ユーザデータ所得
  dim USER, sUN, Utype
  USER   = UCase(Session.Contents("userid"))
  sUN    = Session.Contents("sUN")
  Utype  = Session.Contents("UType")

'データ取得
  dim Mord,CONnum,CMPcd(5),HedId,Rmon,Rday
  dim Hto,CONsize,CONtype,CONhite,CONsitu,CONtear,TrhkSen,MrSk,MaxW
  dim UpFlag,param,i,j,WkContrlNo, ret,ErrerM
  dim SendUser
  ret = true
  Mord   = Request("Mord")
  UpFlag = Request("UpFlag")
  CONnum = "'"& Request("CONnum") &"'"
  For Each param In Request.Form
    If Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next
  Rmon    = Right("00" & Request("Rmon") ,2)
  Rday    = Right("00" & Request("Rday") ,2)
  HedId=Request("HedId")
  HTo=Request("HTo")
  CONsize =Request("CONsize")
  CONtype =Request("CONtype")
  CONhite =Request("CONhite")
  CONsitu =Request("CONsitu")
  CONtear =Request("CONtear")
  TrhkSen =Request("TrhkSen")
  MrSk =Request("MrSk")
  MaxW =Request("MaxW")

'エラートラップ開始
  on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS

'データ整形
  dim FullName,RFlag
  RFlag=0
  FullName= "Null"
  If UpFlag<2 Then
   '元請陸運業者名取得
'      If CMPcd(0) <> "" Then    ' Commented 2003.08.30
      If CMPcd(1) <> "" Then     ' Added 2003.08.30
      StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & CMPcd(1) &"'"
      ObjRS.Open StrSQL, ObjConn
      FullName = "'" & ObjRS("FullName") & "'"
      ObjRS.close
    End If
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB切断
      jampErrerP "1","b202","04","空搬入：データ登録","102","元請陸運業者名取得に失敗<BR>"&StrSQL
    end if
  End If
  If HedId = "" Then
    HedId   = "Null"
  Else
    HedId = "'" & HedId & "'"
  End If

  For i=1 To 4
    If CMPcd(i) = "" Then
      CMPcd(i) = "Null"
    Else
      If CMPcd(i) = Session.Contents("COMPcd") Then
        RFlag=1
      End If
      CMPcd(i) = "'" & CMPcd(i) & "'"
    End If
  Next

  '作業予定日の年度を決定
  If DayTime(1) > Rmon Then	'来年
    YY = DayTime(0) +1
  ElseIf DayTime(1) = Rmon AND DayTime(2) > Rday Then	'CW-043
    YY = DayTime(0) +1					'CW-043
  Else
    YY = DayTime(0)
  End If
  If Rmon = "00" Or Rday = "00" Then
    Yotei= "Null"
  Else
'3th chage      Yotei= "'" & YY &"/"& Rmon &"'"
      Yotei= "'" & YY &"/"& Rmon &"/"& Rday &"'" 
  End If
  If Mord = 0 Then	'初期登録
    '登録重複チェック
    dim dummy
    checkComInfo  ObjConn, ObjRS,CONnum,"2", "1", dummy , ret

    If ret Then
     WriteLogH "b202", "空搬入事前情報入力","02",""
     '作業管理番号採番
      getWkContrlNo ObjConn, ObjRS, sUN, WkContrlNo
     'データ登録
      StrSQL = "Insert Into hITCommonInfo (WkContrlNo,UpdtTime,UpdtPgCd,UpdtTmnl,Status," &_
               "Process,WkType,FullOutType,InPutDate,UpdtUserCode,WkNo,ContNo,ContSize," &_
               "ContType,ContHeight,Material,TareWeight,CustOK,MaxWght," &_
               "RegisterType,RegisterName,RegisterCode,TruckerSubCode1," &_
               "HeadID,WorkDate,TruckerName,Comment1,TruckerSubName1) " &_
               "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"',"&_
               "'0','R','2',Null,'"& Now() &"','"& USER &"',Null,"& CONnum &","&_
               "'"& CONsize &"','"& CONtype &"','"& CONhite &"','"& CONsitu &"','"& CONtear &"'," &_
               "'"& MrSk & "','"& MaxW &"','"& Utype &"','"& sUN &"','"& CMPcd(0) &"',"& CMPcd(1) &","&_
                HedId &","& Yotei &","& FullName &",'"& Request("Comment1") &"','" & Request("TruckerSubName") & "'" & ")"
'C-002 ADD  : ,Comment1 AND ,'"& Request("Comment1") &"'
	 SendUser = CMPcd(1)
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b202","04","空搬入：データ登録","103","SQL:<BR>"&StrSQL
      end if

  '紹介テーブル登録
      StrSQL = "Insert Into hITReference (WkContrlNo, UpdtTime, UpdtPgCd,UpdtTmnl," &_
               "TruckerFlag1,TruckerFlag2,TruckerFlag3,TruckerFlag4)" &_
               "values ('"& WkContrlNo &"','"& Now() &"','PREDEF01','"& USER &"'," &_
               "'"&RFlag&"','0','0','0')"
      ObjConn.Execute(StrSQL)
      if err <> 0 then
        Set ObjRS = Nothing
        jampErrerPDB ObjConn,"1","b202","04","空搬入：データ登録","103","SQL:<BR>"&StrSQL
      end if
    Else
      ErrerM="指定のコンテナは操作中に他者によって登録されました。"
    End If
  Else			'更新
    WriteLogH "b202", "空搬入事前情報入力","14",""
'CW-005	ADD START ↓↓↓↓↓↓↓
   '完了・更新チェック
    If UpFlag <>5 Then
      StrSQL="SELECT ITC.WorkCompleteDate, ITR.TruckerFlag"& UpFlag &" AS Flag "&_
             "FROM hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo "&_
             "Where ContNo="& CONnum &" AND Process='R' AND WkType='2'"
    Else
      StrSQL="SELECT WorkCompleteDate FROM hITCommonInfo " &_
             "Where ContNo="& CONnum &" AND Process='R' AND WkType='2'"
    End If
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      ObjRS.Close
      Set ObjRS = Nothing
      jampErrerPDB ObjConn,"1","b10"&(2+Flag),"14","実搬出：データ登録","101","SQL:<BR>"&StrSQL
    end if
    If NOT IsNull(ObjRS("WorkCompleteDate")) Then 
      ret=false
      ErrerM="指定の作業は画面操作中に作業が完了したため、更新はキャンセルされました。"
    End If
   'チェック
    If UpFlag <>5 Then
      If Trim(ObjRS("Flag"))=1 Then 
        ret=false
        ErrerM="指定の作業は画面操作中に指示先に受諾されたため、更新はキャンセルされました。"
      End If
    End If
    ObjRS.close
    If ret Then
'CW-005	End ADD ↑↑↑↑↑↑↑
      If Mord <> 2  Then	'更新
        dim tmpStr
        If FullName <> "Null" Then
          FullName=",TruckerName="& FullName &" "
        Else
          FullName=" "
        End If
        If UpFlag = 5 Then
          tmpStr = " "
        Else
          tmpStr=" TruckerSubCode"& UpFlag &"="& CMPcd(UpFlag) &","
          SendUser = CMPcd(UpFlag)
        End If
        
        tmpStr= tmpStr & " TruckerSubName"& UpFlag &"='"& Request("TruckerSubName") & "',"
        
      'データ更新
        StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                 "UpdtTmnl='"& USER &"', Status='0', Process='R', " &_
                 "UpdtUserCode='"& USER &"', "& tmpStr &_
                 "HeadID="& HedId &", WorkDate="& Yotei &", ContSize='"& CONsize &"', "&_
                 "ContType='"& CONtype &"', ContHeight='"& CONhite &"', Material='"& CONsitu &"', "&_
                 "TareWeight='"& CONtear &"',CustOK='"& MrSk & "', MaxWght='"& MaxW &"' "& FullName &_
                 ", Comment1='"& Request("Comment1") &"' "&_
                 "Where ContNo="& CONnum &" AND Process='R' AND WkType='2'"
'C-002 ADD This Line : ", Comment1='"& Request("Comment1") &"' "&_
        ObjConn.Execute(StrSQL)
          if err <> 0 then
            Set ObjRS = Nothing
            jampErrerPDB ObjConn,"1","b202","14","空搬入：データ登録","104","SQL:<BR>"&StrSQL
          end if
     '参照フラグ更新
        If UpFlag = 5 Then
          tmpStr = " "
        Else
          If UpFlag = 1 AND Mid(CMPcd(1),2,2) = UCase(Session.Contents("COMPcd")) Then 
            tmpStr = ", TruckerFlag1=1 "
          Else
            tmpStr = ", TruckerFlag"& UpFlag &"=0 "
          End If
        End If
        UpFlag = UpFlag-1
        If UpFlag = 0 Then
          tmpStr = tmpStr&" "
        Else
          tmpStr=tmpStr&", TruckerFlag"& UpFlag &"=1 "
        End If
        StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                 "UpdtTmnl='"& USER &"'"&tmpStr&_
                 "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                 "WHERE ContNo="& CONnum &" AND Process='R' AND WkType='2')"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b202","14","空搬入：データ登録","104","SQL:<BR>"&StrSQL
        end if
      Else	'保留
      'ヘッダID更新
        If UpFlag=5 Then
          tmpStr=""
        Else
          tmpStr=", TruckerSubCode"& UpFlag &"=Null"
        End If
        StrSQL = "UPDATE hITCommonInfo SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                 "UpdtTmnl='"& USER &"', Status='0', Process='R', " &_
                 "UpdtUserCode='"& USER &"'"& tmpStr &", HeadID=Null " &_
                 "Where ContNo="& CONnum &" AND Process='R' AND WkType='2'"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b202","15","空搬入：保留","102","SQL:<BR>"&StrSQL
        end if

       '参照フラグ更新
        StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                 "UpdtTmnl='"& USER &"', TruckerFlag"& UpFlag-1 &"=2 "&_
                 "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                 "WHERE ContNo="& CONnum &" AND Process='R' AND WkType='2')"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b202","15","空搬入：保留","102","SQL:<BR>"&StrSQL
        end if
      End If
    End If		'CW-005
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
		jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","空搬入：メール送信","104","SQL:<BR>"&StrSQL
	end if

	ComInterval = ObjRS("ComInterval")
	ObjRS.Close
		
	if SendUser <> "" then
	''作業発生配信情報の取得
		StrSQL = "SELECT T.*, "
		StrSQL = StrSQL & "CASE WHEN U.NameAbrev IS NULL THEN U.FullName ELSE U.NameAbrev END AS USERNAME "
		StrSQL = StrSQL & "FROM mUsers U, "
		StrSQL = StrSQL & "(SELECT T.* FROM TargetOperation T, mUsers U WHERE T.UserCode = U.UserCode "
		StrSQL = StrSQL & "AND U.HeadCompanyCode =" & SendUser & ") T "
		StrSQL = StrSQL & "WHERE U.UserCode = '" & USER & "'"
		
		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
	'''DB切断
			DisConnDBH ObjConn, ObjRS
			jampErrerPDB ObjConn,"1","b10"&(2+Flag),"16","空搬入：メール送信","104","SQL:<BR>"&StrSQL
		end if

		Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
		Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4, mailFlag5
		Dim mailSubject, mailBody,WorkName
		Dim SendTime, UpdateSendTime
		Dim fp, fobj, tfile
		
' 2009/03/10 R.Shibuta Add-S
	'''SMTPサーバ名の設定
		svName   = "slitdns2.hits-h.com"
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
		if Trim(ObjRS("Email1")) <> "" AND ObjRS("FlagRecEmp1") = "1" then
			mailTo = mailTo & Trim(ObjRS("Email1"))
			mailFlag1 = 1
		else
			mailFlag1 = 0
		end if

		if Trim(ObjRS("Email2")) <> "" AND ObjRS("FlagRecEmp2") = "1" then
			if mailFlag1 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email2"))
			else
				mailTo = mailTo & Trim(ObjRS("Email2"))
			end if
				mailFlag2 = 1
		else
			mailFlag2 = 0
		end if

		if Trim(ObjRS("Email3")) <> "" AND ObjRS("FlagRecEmp3") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
			else
				mailTo = mailTo & Trim(ObjRS("Email3"))
			end if
			mailFlag3 = 1
		else
			mailFlag3 = 0
		end if

		if Trim(ObjRS("Email4")) <> "" AND ObjRS("FlagRecEmp4") = "1" then
			if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
				mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
			else
				mailTo = mailTo & Trim(ObjRS("Email4"))
			end if
			mailFlag4 = 1
		else
			mailFlag4 = 0
		end if

		if Trim(ObjRS("Email5")) <> "" AND ObjRS("FlagRecEmp5") = "1" then
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
		mailBody = "空搬入作業" & "発生 (" & Trim(ObjRS("USERNAME")) & "様より)" & vbCrLf & vbCrLf
		mailBody = mailBody & "空搬入作業" & "が発生しました。" & vbCrLf
		mailBody = mailBody & "詳しくはHiTSの事前情報登録の画面をご参照下さい。"
			
		'メール送信時刻から現在の時刻が通信間隔以上の場合はメールを送信する。
		
		if Trim(mailTo) <> "" Then
			if ObjRS("RecEmpDate") < DateAdd("n",(ComInterval * -1), Now()) OR IsNull(ObjRS("RecEmpDate")) = True then
				rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
				sendTime=Now
			end if

			If rc = "" Then
				'''メール送信日付の更新を行う。
				StrSQL = "UPDATE TargetOperation SET UpdtTime='" & Now() & "', UpdtPgCd='dmi140',"
				StrSQL = StrSQL & " UpdtTmnl='" & USER & "',"&  "RecEmpDate='" & Now() & "'"
				StrSQL = StrSQL &"WHERE UserCode = '" & Trim(ObjRS("UserCode")) & "'"

				ObjConn.Execute(StrSQL)
				if err <> 0 then
					Set ObjRS = Nothing
					jumpErrorPDB ObjConn,"1","c104","14","空搬入：メール送信","104","SQL:<BR>"&StrSQL
				end if
			else
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
' 2009/03/10 R.Shibuta Add-E
	end if
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前空搬入登録・更新</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------事前空搬入登録・更新--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR><TD align=center>
<% If ret Then%>
  <% If Mord=0 Then %>
   登録しました。<BR>画面は自動的に閉じられます。
    <SCRIPT language=JavaScript>
      try{
        window.opener.parent.List.location.href="./dmo110F.asp"
      }catch(e){}
      window.close();
    </SCRIPT>
  <% Else %>
   更新しました。<BR>画面は自動的に閉じられます。
    <SCRIPT language=JavaScript>
      try{
        window.opener.parent.DList.location.href="./dmo110L.asp"
        window.opener.parent.Top.location.href="./dmo110T.asp"
      }catch(e){}
      window.close();
    </SCRIPT>
  <% End If %>
<% Else %>
   <DIV class=alert><%=ErrerM%></DIV><BR>
   <INPUT type=button value="閉じる" onClick="window.close()">
<% End If %>
  </TD>
  </TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
