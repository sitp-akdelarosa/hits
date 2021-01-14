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
	'''Microsoft ADO用のadovbs.incにて提供されている
	Const adBoolean = 11
	Const adDBTimeStamp = 135
	Const adInteger = 3
	Const adChar = 129
	Const adParamInput = &H0001
	Const adParamReturnValue = &H0004

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



	Dim ArrayContNo, ETA, TA, InTime, OLTICDate, DOStatus
	Dim PreDelPermitFlag, DelPermitDate, DemurrageFreeTime
	Dim CYDelTime, DetentionFreeTime, ReturnTime
	Dim VslCode, VoyCtrl
	Dim sp, p0, p1, p2, p3, p4
	Dim strchkNow, strchkOLTDateFrom, strchkOLTDateTo
	Dim PreTsukanFlag


	'''指定コンテナ番号,ＢＬ番号の存在チェック
	'''ＣＹ搬出済でも１０日以内であれば登録可とする(2004.2.16仕様変更)。
	'''ＢＬ番号指定の場合、ＣＹ搬出されていないか、搬出されていても１０日以内のコンテナのみ登録対象とする(2004.2.16仕様変更)。
	Dim Num, Num2, i

	if KIND = 1 then	'''コンテナ番号指定
'		StrSQL = "SELECT ContNo, BLNo FROM ImportCont "
		StrSQL = "SELECT IC.*, VP.ETA, VP.TA "
		StrSQL = StrSQL & " FROM ImportCont IC, VslPort VP "
		StrSQL = StrSQL & " WHERE IC.ContNo='"& NUMBER &"'"
		StrSQL = StrSQL & " AND IC.UpdtTime = (SELECT max(UpdtTime) FROM ImportCont WHERE ContNo='"& NUMBER &"') "
		StrSQL = StrSQL & " AND IC.VslCode = VP.VslCode "
		StrSQL = StrSQL & " AND IC.VoyCtrl = VP.VoyCtrl "
		StrSQL = StrSQL & " AND VP.PortCode ='JPHKT' "
		StrSQL = StrSQL & " AND (IC.CYDelTime is NULL "
		StrSQL = StrSQL & " OR (IC.CYDelTime is not NULL "
		StrSQL = StrSQL & " AND DATEDIFF(d,IC.CYDelTime,GETDATE()) >= 0 "
		StrSQL = StrSQL & " AND DATEDIFF(d,IC.CYDelTime,GETDATE()) <= 10)) "
	elseif KIND = 2 then	'''ＢＬ番号指定
		StrSQL = "SELECT BL.BLNo, IC.*, VP.ETA, VP.TA "
		StrSQL = StrSQL & " FROM BL, ImportCont IC, VslPort VP "
		StrSQL = StrSQL & " WHERE BL.BLNo='"& NUMBER &"'"
		StrSQL = StrSQL & " AND BL.VslCode = IC.VslCode "
		StrSQL = StrSQL & " AND BL.VoyCtrl = IC.VoyCtrl "
		StrSQL = StrSQL & " AND BL.BLNo = IC.BLNo "
		StrSQL = StrSQL & " AND BL.UpdtTime = (SELECT max(BL.UpdtTime) FROM BL WHERE BL.BLNo='"& NUMBER &"') "
		StrSQL = StrSQL & " AND IC.VslCode = VP.VslCode "
		StrSQL = StrSQL & " AND IC.VoyCtrl = VP.VoyCtrl "
		StrSQL = StrSQL & " AND VP.PortCode ='JPHKT' "
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


	if Num > 0 then
		'''搬出可否判定用ストアードプロシジャの呼び出しのための設定
		set sp = Server.CreateObject("ADODB.Command")
		set sp.ActiveConnection = ObjConn
		sp.CommandText = "{?=call DelPermitCheck(?,?,?)}"
		Set p0 = sp.CreateParameter("ret", adBoolean, adParamReturnValue)
		sp.Parameters.Append p0
		Set p1 = sp.CreateParameter("VslCode", adChar, adParamInput, 7)
		sp.Parameters.Append p1
		Set p2 = sp.CreateParameter("VoyCtrl", adInteger, adParamInput)
		sp.Parameters.Append p2
		Set p3 = sp.CreateParameter("ContNo", adChar, adParamInput, 12)
		sp.Parameters.Append p3


		if KIND = 1 then		'''コンテナ番号指定の場合
			ReDim ArrayContNo(1), ETA(1), TA(1), InTime(1), OLTICDate(1), DOStatus(1)
			ReDim PreDelPermitFlag(1), DelPermitDate(1), DemurrageFreeTime(1)
			ReDim CYDelTime(1), DetentionFreeTime(1), ReturnTime(1)
			ReDim VslCode(1), VoyCtrl(1)
			ReDim strchkNow(1), strchkOLTDateFrom(1), strchkOLTDateTo(1)
			ReDim PreTsukanFlag(1)

			'''TargetContainersテーブルへの初期値設定のためのデータ取り出し
			ArrayContNo(0) = NUMBER
			ETA(0) = ObjRS("ETA")
			TA(0) = ObjRS("TA")
			InTime(0) = ObjRS("InTime")
			OLTICDate(0) = ObjRS("OLTICDate")
			DOStatus(0) = ObjRS("DOStatus")
			DelPermitDate(0) = ObjRS("DelPermitDate")
			if IsNull(ObjRS("FreeTimeExt")) then
				DemurrageFreeTime(0) = ObjRS("FreeTime")
			else
				DemurrageFreeTime(0) = ObjRS("FreeTimeExt")
			end if
			CYDelTime(0) = ObjRS("CYDelTime")
			DetentionFreeTime(0) = ObjRS("DetentionFreeTime")
			ReturnTime(0) = ObjRS("ReturnTime")
			VslCode(0) = ObjRS("VslCode")
			VoyCtrl(0) = ObjRS("VoyCtrl")

			'''搬出可否判定
			'''ImportContテーブルのVslCode, VoyCtrl, ContNoが同じでBLNoだけが異なるレコードが存在する場合、
			'''当該レコードについても条件をクリアできているかチェックする。
			sp("VslCode") = VslCode(0)
			sp("VoyCtrl") = VoyCtrl(0)
			sp("ContNo") = ArrayContNo(0)
			'''ストアードプロシジャの呼び出し
			sp.Execute
			if sp("ret") = True then 
				PreDelPermitFlag(0) = "Y"
			else
				PreDelPermitFlag(0) = "N"
			end if

			'''通関○×判定		Added 20040331
			strchkNow(0) = DispDateTime(Now,8)
			strchkOLTDateFrom(0) = DispDateTime(ObjRS("OLTDateFrom"),8)
			strchkOLTDateTo(0) = DispDateTime(ObjRS("OLTDateTo"),8)
			PreTsukanFlag(0) = 0
			if Trim(ObjRS("OLTICFlag"))="I" then
				if Trim(ObjRS("OLTICNo"))<>"" then
					PreTsukanFlag(0) = "Y"
				else
					PreTsukanFlag(0) = "N"
				end if
			else
				if strchkNow(0) >= strchkOLTDateFrom(0) and strchkNow(0) <= strchkOLTDateTo(0) then
					PreTsukanFlag(0) = "Y"
				else
					PreTsukanFlag(0) = "N"
				end if
			end if
			''' 搬出されていたら○とする
			if DispDateTime(ObjRS("CYDelTime"),0)<>"" then
				PreTsukanFlag(0) = "Y"
			end if
			'''通関○×判定		Added 20040331 ここまで

		end if

		if KIND=2 then		'''ＢＬ番号指定の場合、紐付いているコンテナ番号を変数に格納
			ReDim ArrayContNo(Num), ETA(Num), TA(Num), InTime(Num), OLTICDate(Num), DOStatus(Num)
			ReDim PreDelPermitFlag(Num), DelPermitDate(Num), DemurrageFreeTime(Num)
			ReDim CYDelTime(Num), DetentionFreeTime(Num), ReturnTime(Num)
			ReDim VslCode(Num), VoyCtrl(Num)
			ReDim strchkNow(Num), strchkOLTDateFrom(Num), strchkOLTDateTo(Num)
			ReDim PreTsukanFlag(Num)

			'''TargetContainersテーブルへの初期値設定のためのデータ取り出し
			for i=0 to Num-1
				ArrayContNo(i) = ObjRS("ContNo")
				ETA(i) = ObjRS("ETA")
				TA(i) = ObjRS("TA")
				InTime(i) = ObjRS("InTime")
				OLTICDate(i) = ObjRS("OLTICDate")
				DOStatus(i) = ObjRS("DOStatus")
				DelPermitDate(i) = ObjRS("DelPermitDate")
				if IsNull(ObjRS("FreeTimeExt")) then
					DemurrageFreeTime(i) = ObjRS("FreeTime")	''' BUG Modified 20040419 (0->i)
				else
					DemurrageFreeTime(i) = ObjRS("FreeTimeExt")	''' BUG Modified 20040419 (0->i)
				end if
				CYDelTime(i) = ObjRS("CYDelTime")
				DetentionFreeTime(i) = ObjRS("DetentionFreeTime")
				ReturnTime(i) = ObjRS("ReturnTime")
				VslCode(i) = ObjRS("VslCode")
				VoyCtrl(i) = ObjRS("VoyCtrl")

				'''搬出可否判定
				'''ImportContテーブルのVslCode, VoyCtrl, ContNoが同じでBLNoだけが異なるレコードが存在する場合、
				'''当該レコードについても条件をクリアできているかチェックする。
				sp("VslCode") = VslCode(i)
				sp("VoyCtrl") = VoyCtrl(i)
				sp("ContNo") = ArrayContNo(i)
				'''ストアードプロシジャの呼び出し
				sp.Execute
				if sp("ret") = True then 
					PreDelPermitFlag(i) = "Y"
				else
					PreDelPermitFlag(i) = "N"
				end if

				'''通関○×判定		Added 20040331
				strchkNow(i) = DispDateTime(Now,8)
				strchkOLTDateFrom(i) = DispDateTime(ObjRS("OLTDateFrom"),8)
				strchkOLTDateTo(i) = DispDateTime(ObjRS("OLTDateTo"),8)
				PreTsukanFlag(i) = 0
				if Trim(ObjRS("OLTICFlag"))="I" then
					if Trim(ObjRS("OLTICNo"))<>"" then
						PreTsukanFlag(i) = "Y"
					else
						PreTsukanFlag(i) = "N"
					end if
				else
					if strchkNow(i) >= strchkOLTDateFrom(i) and strchkNow(i) <= strchkOLTDateTo(i) then
						PreTsukanFlag(i) = "Y"
					else
						PreTsukanFlag(i) = "N"
					end if
				end if
				''' 搬出されていたら○とする
				if DispDateTime(ObjRS("CYDelTime"),0)<>"" then
					PreTsukanFlag(i) = "Y"
				end if
				'''通関○×判定		Added 20040331 ここまで

				ObjRS.MoveNext

			next
		end if
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
				StrSQL =  StrSQL & "ETA, TA, InTime, "
				StrSQL =  StrSQL & "ListDate, DOStatus, PreDelPermitFlag, "
				StrSQL =  StrSQL & "DelPermitDate, DemurrageFreeTime, "
				StrSQL =  StrSQL & "CYDelTime, DetentionFreeTime, ReturnTime, PreTsukanFlag) "
				StrSQL =  StrSQL & "values ('" & USER & "','" & DayTime & "','STATUS01','" & USER & "','" & DayTime & "','R',"
				StrSQL =  StrSQL & "'" & NUMBER & "',Null, Null,"
				StrSQL =  StrSQL & "'0','0','0','0','0','0',"
				StrSQL =  StrSQL & "'0','0','0','0',"
				if IsNull(ETA(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & ETA(0) & "',"
				end if
				if IsNull(TA(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & TA(0) & "',"
				end if
				if IsNull(InTime(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & InTime(0) & "',"
				end if
				if IsNull(OLTICDate(0)) then		'''テーブルのフィールド名称はListDateのまま変更していない	Changed 20040331
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & OLTICDate(0) & "',"
				end if
				if IsNull(DOStatus(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & DOStatus(0) & "',"
				end if
				if IsNull(PreDelPermitFlag(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & PreDelPermitFlag(0) & "',"
				end if
				if IsNull(DelPermitDate(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & DelPermitDate(0) & "',"
				end if
				if IsNull(DemurrageFreeTime(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & DemurrageFreeTime(0) & "',"
				end if
				if IsNull(CYDelTime(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & CYDelTime(0) & "',"
				end if
				if IsNull(DetentionFreeTime(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & Trim(DetentionFreeTime(0)) & "',"
				end if
				if IsNull(ReturnTime(0)) then
					StrSQL =  StrSQL & "NULL,"
				else
					StrSQL =  StrSQL & "'" & ReturnTime(0) & "',"
				end if
				if IsNull(PreTsukanFlag(0)) then		'''Added 20040331
					StrSQL =  StrSQL & "NULL)"
				else
					StrSQL =  StrSQL & "'" & PreTsukanFlag(0) & "')"
				end if


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
					StrSQL =  StrSQL & "CYDelTime, DetentionFreeTime, ReturnTime, PreTsukanFlag) "
					StrSQL =  StrSQL & "values ('" & USER & "','" & DayTime & "','STATUS01','" & USER & "','" & DayTime & "','R',"
					StrSQL =  StrSQL & "'" & ArrayContNo(i) & "', '" & NUMBER & "', Null,"
					StrSQL =  StrSQL & "'0','0','0','0','0','0',"
					StrSQL =  StrSQL & "'0','0','0','0',"
					if IsNull(ETA(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & ETA(i) & "',"
					end if
					if IsNull(TA(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & TA(i) & "',"
					end if
					if IsNull(InTime(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & InTime(i) & "',"
					end if
					if IsNull(OLTICDate(i)) then		'''テーブルのフィールド名称はListDateのまま変更していない	Changed 20040331
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & OLTICDate(i) & "',"
					end if
					if IsNull(DOStatus(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
					StrSQL =  StrSQL & "'" & DOStatus(i) & "',"
					end if
					if IsNull(PreDelPermitFlag(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & PreDelPermitFlag(i) & "',"
					end if
					if IsNull(DelPermitDate(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & DelPermitDate(i) & "',"
					end if
					if IsNull(DemurrageFreeTime(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & DemurrageFreeTime(i) & "',"
					end if
					if IsNull(CYDelTime(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & CYDelTime(i) & "',"
					end if
					if IsNull(DetentionFreeTime(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & Trim(DetentionFreeTime(i)) & "',"
					end if
					if IsNull(ReturnTime(i)) then
						StrSQL =  StrSQL & "NULL,"
					else
						StrSQL =  StrSQL & "'" & ReturnTime(i) & "',"
					end if
					if IsNull(PreTsukanFlag(i)) then		'''Added 20040331
						StrSQL =  StrSQL & "NULL)"
					else
						StrSQL =  StrSQL & "'" & PreTsukanFlag(i) & "')"
					end if

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
