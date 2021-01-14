<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst201.asp				_/
'_/	Function	:ステータス配信依頼新規登録			_/
'_/	Date			:2004/01/15				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''セッションの有効性をチェック
	CheckLoginH

	'''サーバ日付の取得
	Dim DayTime
	DayTime = Now()

	'''登録しました画面にて「最新の情報に更新」でSubmitされた場合の対策
	if Session.Contents("InsertSubmitted")="False" then

	'''データ取得
	Dim USER, KIND, NUMBER, ErrCode
	USER   = UCase(Session.Contents("userid"))
	KIND = Request.Form("ContORBL")
	NUMBER = Request.Form("ContBLNo")
	ErrCode = 0

	'''エラートラップ開始
	on error resume next
	'''DB接続
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	'''指定コンテナ番号,ＢＬ番号の存在チェック
	'''ＣＹ搬出済でも１０日以内であれば登録可とする(2004.2.16仕様変更)。
	'''ＢＬ番号指定の場合、ＣＹ搬出されていないか、搬出されていても１０日以内のコンテナのみ登録対象とする(2004.2.16仕様変更)。
	Dim Num, Num2, ArrayContNo, i
	if KIND = 1 then	'''コンテナ番号指定
		StrSQL = "SELECT ContNo, BLNo FROM ImportCont "
		StrSQL = StrSQL & " WHERE ContNo='"& NUMBER &"'"
		StrSQL = StrSQL & " AND UpdtTime = (SELECT max(UpdtTime) FROM ImportCont WHERE ContNo='"& NUMBER &"') "
		StrSQL = StrSQL & " AND (CYDelTime is NULL "
		StrSQL = StrSQL & " OR (CYDelTime is not NULL "
		StrSQL = StrSQL & " AND DATEDIFF(d,CYDelTime,GETDATE()) >= 0 "
		StrSQL = StrSQL & " AND DATEDIFF(d,CYDelTime,GETDATE()) <= 10)) "
	elseif KIND = 2 then	'''ＢＬ番号指定
		StrSQL = "SELECT BL.BLNo, IC.ContNo FROM BL, ImportCont IC "
		StrSQL = StrSQL & " WHERE BL.BLNo='"& NUMBER &"'"
		StrSQL = StrSQL & " AND BL.VslCode = IC.VslCode "
		StrSQL = StrSQL & " AND BL.VoyCtrl = IC.VoyCtrl "
		StrSQL = StrSQL & " AND BL.BLNo = IC.BLNo "
		StrSQL = StrSQL & " AND BL.UpdtTime = (SELECT max(BL.UpdtTime) FROM BL WHERE BL.BLNo='"& NUMBER &"') "
		StrSQL = StrSQL & " AND (IC.CYDelTime is NULL "
		StrSQL = StrSQL & " OR (IC.CYDelTime is not NULL "
		StrSQL = StrSQL & " AND DATEDIFF(d,IC.CYDelTime,GETDATE()) >= 0 "
		StrSQL = StrSQL & " AND DATEDIFF(d,IC.CYDelTime,GETDATE()) <= 10)) "
	else
		response.write("KIND error!")
		response.end
	end if

	ObjRS.Open StrSQL, ObjConn, 3, 1
	if err <> 0 then
		'''DB切断
		DisConnDBH ObjConn, ObjRS
		jumpErrorP "1","c102","01","ステータス配信依頼新規登録","101","SQL:<BR>"&strSQL
	end if
	Num = ObjRS.RecordCount

	if KIND=2 then		'''ＢＬ番号指定の場合、紐付いているコンテナ番号を変数に格納
		ReDim ArrayContNo(Num)
		for i=0 to Num-1
			ArrayContNo(i) = ObjRS("ContNo")
			ObjRS.MoveNext
		next
	end if

	if KIND=1 then
		if Num > 0 then
			if Trim(ObjRS("BLNo")) = "" then
				'''コンテナ番号はセットされているがＢＬ番号がセットされていないレコードが指定された場合
				Response.Write("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>")
				Response.Write("<HTML>")
				Response.Write("<HEAD>")
				Response.Write("<LINK REL='stylesheet' TYPE='text/css' HREF='./style.css'>")
				Response.Write("<TITLE>ステータス配信依頼新規登録</TITLE>")
				Response.Write("<META content='text/html; charset=Shift_JIS' http-equiv=Content-Type>")
				Response.Write("<SCRIPT Language='JavaScript'>")
				Response.Write("<!--")
				Response.Write("function CloseWin(){")
				Response.Write("try{")
				Response.Write("window.opener.parent.List.location.href='sst100F.asp'")
				Response.Write("}catch(e){}")
				Response.Write("window.close();")
				Response.Write("}")
				Response.Write("// -->")
				Response.Write("</SCRIPT>")
				Response.Write("<META content='MSHTML 5.00.2919.6307' name=GENERATOR></HEAD>")
				Response.Write("<BODY leftMargin='0' topMargin='0' marginheight='0' marginwidth='0'>")
				Response.Write("<TABLE border='0' cellPadding='5' cellSpacing='0' width='100%'>")
				Response.Write("<FORM name='sst201'>")
				Response.Write("<TR><TD>　</TD></TR>")
				Response.Write("<TR>")
				Response.Write("<TD align='center'>")
				Response.Write("指定されたコンテナのＢＬ番号が輸入コンテナテーブルに<BR>設定されていません。<BR><BR><BR>")
				Response.Write("<INPUT type='button' value='閉じる' onClick='javascript:window.close();'>")
				Response.Write("</TD>")
				Response.Write("</TR>")
				Response.Write("</FORM>")
				Response.Write("</TABLE>")
				Response.Write("</BODY>")
				Response.Write("</HTML>")
				ObjRS.close
				Response.end
			end if
		end if
	end if
	ObjRS.close


	if Num > 0 then  '''指定されたコンテナ番号またはＢＬ番号が輸入コンテナテーブルに存在する場合
		'''指定されたコンテナ番号またはＢＬ番号を同じユーザがすでに指定しているかどうかのチェック
		if KIND = 1 then '''コンテナ番号指定
			StrSQL = "SELECT ContNo FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' AND ContNo='" & NUMBER & "'"
			StrSQL = StrSQL & " AND BLNo is NULL"
		elseif KIND = 2 then	'''ＢＬ番号指定の場合
			StrSQL = "SELECT BLNo FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' AND BLNo='" & NUMBER & "'"
		else
			response.write("KIND error!")
			response.end
		end if
		ObjRS.Open StrSQL, ObjConn, 3, 1
		if err <> 0 then
			'''DB切断
			DisConnDBH ObjConn, ObjRS
			jumpErrorP "2","c102","01","ステータス配信依頼新規登録","101","SQL:<BR>"&strSQL
		end if
		Num2 = ObjRS.RecordCount
		ObjRS.close

		if Num2 > 0 then		'''すでに同じユーザが同じコンテナ番号、ＢＬ番号を登録している
			ErrCode = 1
		else

		'''データ登録
			if KIND = 1 then		''''コンテナ番号指定の場合
				StrSQL = "INSERT INTO TargetContainers (UserCode, UpdtTime, UpdtPgCd, UpdtTmnl, RegisterDate, Process, "
				StrSQL =  StrSQL & "ContNo, BLNo, LatestSentTime, "
				StrSQL =  StrSQL & "FlagETA, FlagTA, FlagInTime, FlagList, FlagDOStatus, FlagDelPermit, "
				StrSQL =  StrSQL & "FlagDemurrageFreeTime, FlagCYDelTime, FlagDetentionFreeTime, FlagReturnTime, "
				StrSQL =  StrSQL & "ETA, TA, InTime, ListDate, DOStatus, PreDelPermitFlag, DelPermitDate, DemurrageFreeTime, "
				StrSQL =  StrSQL & "CYDelTime, DetentionFreeTime, ReturnTime) "
				StrSQL =  StrSQL & "values ('" & USER & "','" & DayTime & "','STATUS01','" & USER & "','" & DayTime & "','R',"
				StrSQL =  StrSQL & "'" & NUMBER & "',Null, Null,"
				StrSQL =  StrSQL & "'0','0','0','0','0','0',"
				StrSQL =  StrSQL & "'0','0','0','0',"
				StrSQL =  StrSQL & "Null,Null,Null,Null,Null,'N',Null,Null,"
				StrSQL =  StrSQL & "Null,Null,Null)"
				ObjConn.Execute(StrSQL)
				if err <> 0 then
					Set ObjRS = Nothing
					jumpErrorPDB ObjConn,"1","c102","01","ステータス配信依頼新規登録","103","SQL:<BR>"&StrSQL
				end if

			elseif KIND = 2 then
				for i=0 to Num-1
					StrSQL = "INSERT INTO TargetContainers (UserCode, UpdtTime, UpdtPgCd, UpdtTmnl, RegisterDate, Process, "
					StrSQL =  StrSQL & "ContNo, BLNo, LatestSentTime, "
					StrSQL =  StrSQL & "FlagETA, FlagTA, FlagInTime, FlagList, FlagDOStatus, FlagDelPermit, "
					StrSQL =  StrSQL & "FlagDemurrageFreeTime, FlagCYDelTime, FlagDetentionFreeTime, FlagReturnTime, "
					StrSQL =  StrSQL & "ETA, TA, InTime, ListDate, DOStatus, PreDelPermitFlag, DelPermitDate, DemurrageFreeTime, "
					StrSQL =  StrSQL & "CYDelTime, DetentionFreeTime, ReturnTime) "
					StrSQL =  StrSQL & "values ('" & USER & "','" & DayTime & "','STATUS01','" & USER & "','" & DayTime & "','R',"
					StrSQL =  StrSQL & "'" & ArrayContNo(i) & "', '" & NUMBER & "', Null,"
					StrSQL =  StrSQL & "'0','0','0','0','0','0',"
					StrSQL =  StrSQL & "'0','0','0','0',"
					StrSQL =  StrSQL & "Null,Null,Null,Null,Null,'N',Null,Null,"
					StrSQL =  StrSQL & "Null,Null,Null)"

					ObjConn.Execute(StrSQL)
					if err <> 0 then
						Set ObjRS = Nothing
						jumpErrorPDB ObjConn,"1","c102","01","ステータス配信依頼新規登録","103","SQL:<BR>"&StrSQL
					end if
				next
			else
				response.write("KIND error!")
				response.end
			end if

			'''ログ出力
			WriteLogH "c102", "ステータス配信依頼新規登録","01",""
			ObjRS.close
		end if

	else		'''指定されたコンテナ番号、ＢＬ番号が存在しない
		ErrCode = 9
	end if

	'''DB接続解除
	DisConnDBH ObjConn, ObjRS
	'''エラートラップ解除
	on error goto 0

	Session.Contents("InsertSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信依頼新規登録</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="sst100F.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------ステータス配信依頼新規登録--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst201">
	<TR><TD>　</TD></TR>
<% if ErrCode=0 then %>
	<TR>
		<TD align="center">
			登録しました。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
<% elseif ErrCode=1 then %>
	<TR>
		<TD align="center">
			指定されたコンテナ番号またはＢＬ番号はすでに登録されています。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="javascript:window.close();">
		</TD>
	</TR>
<% elseif ErrCode=9 then %>
	<TR>
		<TD align="center">
			指定されたコンテナ番号またはＢＬ番号は存在しないか、<BR>
			搬出後１１日以上経過しているため登録できません。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="javascript:window.close();">
		</TD>
	</TR>
<% end if %>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("InsertSubmitted")="False"のelse処理 %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>ステータス配信依頼新規登録</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="sst100F.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------ステータス配信依頼新規登録--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="sst201">
	<TR><TD>　</TD></TR>
	<TR>
		<TD align="center">
			登録はすでに完了しています。<BR><BR><BR>
			<INPUT type="button" value="閉じる" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("InsertSubmitted")="False"のendif処理 %>
<% end if %>
