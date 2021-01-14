<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										   _/
'_/	FileName	:dmi500.asp									   _/
'_/	Function	:作業発生mail即時送信						   _/
'_/	Date		:2009/03/11									   _/
'_/	Code By		:Shibuta									   _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''Microsoft ADO用のadovbs.incにて提供されている
	Const adBoolean = 11
	Const adDBTimeStamp = 135
	Const adInteger = 3
	Const adChar = 129
	Const adParamInput = &H0001
	Const adParamReturnValue = &H0004
	Dim ErrCode
	
	ErrCode = 0
	
	'''セッションの有効性をチェック
	CheckLoginH
	Session.Contents("SendMailSubmitted") = "False"
	'''送信しました画面にて「最新の情報に更新」でSubmitされた場合の対策
	if Session.Contents("SendMailSubmitted") = "False" then

		'''データ取得
		Dim USER, CALLPG, SENDUSER
		Dim Email1, Email2, Email3, Email4, Email5
		Dim UserName,ComInterval,rc
		
		USER = Session.Contents("userid")
		CALLPG = Session.Contents("callpg")
		SENDUSER = Session.Contents("senduser")

		'''DB接続
		Dim ObjConn, ObjRS, StrSQL
		ConnDBH ObjConn, ObjRS

		'''通信間隔取得
		StrSQL = "SELECT ComInterval FROM mParam "

		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			'''DB切断
			DisConnDBH ObjConn, ObjRS
			jumpErrorP "2","c104","01","作業発生mail即時送信","101","SQL:<BR>"&strSQL
		end if

		ComInterval = ObjRS("ComInterval")
		ObjRS.Close

		if SENDUSER <> "" then
		''作業発生配信情報の取得
			StrSQL = "SELECT T.*, "
			StrSQL = StrSQL & "CASE WHEN U.NameAbrev IS NULL THEN U.FullName ELSE U.NameAbrev END AS USERNAME "
			StrSQL = StrSQL & "FROM mUsers U, "
			StrSQL = StrSQL & "(SELECT T.* FROM TargetOperation T, mUsers U WHERE T.UserCode = U.UserCode "
			StrSQL = StrSQL & "AND U.HeadCompanyCode =" & SENDUSER & ") T "
			StrSQL = StrSQL & "WHERE U.UserCode = '" & USER & "'"
			
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
			'''DB切断
				DisConnDBH ObjConn, ObjRS
				jumpErrorP "2","c104","01","作業発生mail即時送信","101","SQL:<BR>"&strSQL
			end if
			'ObjRS.close

			Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
			Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4, mailFlag5
			Dim mailSubject, mailBody,WorkName
			Dim SendTime, UpdateSendTime
		
			'''SMTPサーバ名の設定
			svName   = "slitdns2.hits-h.com"
			'svName = "192.168.17.61"
			attachedFiles = ""
			mailFlag1 = 0
			mailFlag2 = 0
			mailFlag3 = 0
			mailFlag4 = 0
			mailFlag5 = 0
			'''メール送信元アドレスの設定
			mailFrom = "mrhits@hits-h.com"
'			mailFrom = "test@192.168.17.61"
			mailTo = ""

		Select Case CALLPG
			'''実搬出作業
			Case "dmi040"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagDelResults1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if

				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagDelResults2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagDelResults3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagDelResults4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagDelResults5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "実搬出作業"
				SendTime = ObjRS("DelResultsDate")
				UpdateSendTime = "DelResultsDate"

			'''空搬入作業
			Case "dmi140"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagRecEmp1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if
				
				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagRecEmp2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagRecEmp3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagRecEmp4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagRecEmp5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "空搬入作業"
				SendTime = ObjRS("RecEmpDate")
				UpdateSendTime = "RecEmpDate"

			'''実搬入作業
			Case "dmi340"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagRecResults1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if

				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagRecResults2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagRecResults3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagRecResults4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagRecResults5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "実搬入作業"
				SendTime = ObjRS("RecResultsDate")
				UpdateSendTime = "RecResultsDate"

			'''空搬出作業
			Case "dmi240"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagDelEmp1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if

				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagDelEmp2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagDelEmp3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagDelEmp4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagDelEmp5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "空搬出作業"
				SendTime = ObjRS("DelEmpDate")
				UpdateSendTime = "DelEmpDate"
			End Select
			
			Set ObjMail = Server.CreateObject("BASP21")

			mailSubject = "HiTS 作業依頼"
			mailBody = WorkName & "発生 (" & Trim(ObjRS("USERNAME")) & "様より)" & vbCrLf & vbCrLf
			mailBody = mailBody & " (" & Trim(ObjRS("USERNAME")) & "様より)" & vbCrLf & vbCrLf
			mailBody = mailBody & WorkName & "が発生しました。" & vbCrLf
			mailBody = mailBody & "詳しくはHiTSの事前情報登録の画面をご参照下さい。"

			'メール送信時刻から現在の時刻が通信間隔以上の場合はメールを送信する。
			if Trim(mailTo) <> "" Then
				WriteLogH "c104", "作業発生mail即時送信","svName",svName
				WriteLogH "c104", "作業発生mail即時送信","mailTo",mailTo
				WriteLogH "c104", "作業発生mail即時送信","mailFrom",mailFrom
				WriteLogH "c104", "作業発生mail即時送信","mailSubject",mailSubject
				WriteLogH "c104", "作業発生mail即時送信","mailBody",mailBody

				if ComInterval < DateDiff("n",SendTime,Now) then
'					rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
					sendTime=Now
				else
					ErrCode = 8
				end if

				If rc = 0 Then
					'''メール送信日付の更新を行う。
					StrSQL = "UPDATE TargetOperation SET UpdtTime='" & Now() & "', UpdtPgCd='dmi500',"
					StrSQL = StrSQL & " UpdtTmnl='" & USER & "',"&  UpdateSendTime & "='" & Now() & "'"
					StrSQL = StrSQL &"WHERE UserCode = '" & Trim(ObjRS("UserCode")) & "'"

					ObjConn.Execute(StrSQL)
					if err <> 0 then
						Set ObjRS = Nothing
						jumpErrorPDB ObjConn,"1","c104","14","作業発生mail即時送信","104","SQL:<BR>"&StrSQL
					end if
	
					'''ログ出力
					WriteLogH "c104", "作業発生mail即時送信","01",""
					ErrCode = 0
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
							ErrCode = 8
						end if
				end if
			else
				ErrCode = 1
			end if
		end if

		'''DB接続解除
		DisConnDBH ObjConn, ObjRS
		'''エラートラップ解除
		on error goto 0

		Session.Contents("SendMailSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>作業発生mail即時送信</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
<!--	window.close(); -->
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------ステータス配信mail即時送信結果画面--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="dmi500" method="POST">
	<TR><TD>　</TD></TR>
<% if ErrCode=0 then %>
	<TR>
		<TD align="center">
			メール送信しました。<BR>
		</TD>
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
<% elseif ErrCode=1 then %>
	<TR>
		<TD align="center">
			メール送信先が設定されていません。<BR>「設定」メニューにてメールアドレスを登録してください。<BR>
		</TD>
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="window.close()">
		</TD>
	</TR>
<% elseif ErrCode=8 then %>
	<TR>
		<TD align="center">
			メール送信に失敗しました。<BR>
		</TD>
	</TR>
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="window.close()">
		</TD>
	</TR>
<% end if %>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>

<%''' if Session.Contents("SendMailSubmitted") = "False"のelse処理 %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>作業発生mail即時送信</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------作業発生mail即時送信結果画面--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="dmi500" method="POST">
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			処理は既に完了しています。<BR><BR><BR>
		</TD>
	</TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("SendMailSubmitted") = "False"のendif処理 %>
<% end if %>
