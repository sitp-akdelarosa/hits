<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst100L.asp				_/
'_/	Function	:ステータス配信依頼中一覧画面リスト出力		_/
'_/	Date			:2003/12/25				_/
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

	'''ユーザデータ所得
	Dim USER
	USER = UCase(Session.Contents("userid"))

	'''検索からの呼び出しかそうでないかの判定
	Dim SortFlag, SortKey
	if Request("SortFlag") = "" then
		SortFlag = 0
	else
		SortFlag = Request("SortFlag")
		SortKey = Request("SortKey")
	end if

	Dim Num, DtTbl, i, j

	'''DB接続
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

	Select Case SortFlag
		Case "0"		'''初期表示
			'''データ取得
			StrSQL = "SELECT RegisterDate, ContNo, BLNo, LatestSentTime FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' and BLNo is NULL "
			StrSQL = StrSQL & " union "
			StrSQL = StrSQL & " SELECT MAX(RegisterDate), NULL ContNo, BLNo, MAX(LatestSentTime) FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' and BLNo is not NULL "
			StrSQL = StrSQL & " Group BY BLNo"
			StrSQL = StrSQL & " ORDER BY 1 DESC"

			ObjRS.Open StrSQL, ObjConn, 3, 1
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jumpErrorP "1","c101","01","ステータス配信依頼中一覧","101","SQL:<BR>"&strSQL
			end if
			'''対象件数設定
			Num = ObjRS.RecordCount

		Case "2"		'''コンテナ番号指定で検索
			'''データ取得
			StrSQL = "SELECT RegisterDate, ContNo, BLNo, LatestSentTime FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' AND ContNo like '%" & SortKey & "'"
			StrSQL = StrSQL & " AND BLNo is NULL "
			StrSQL = StrSQL & " ORDER BY RegisterDate DESC"

			ObjRS.Open StrSQL, ObjConn, 3, 1
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jumpErrorP "1","c101","01","ステータス配信依頼中一覧","101","SQL:<BR>"&strSQL
			end if
			'''対象件数設定
			Num = ObjRS.RecordCount

		Case "3"		'''ＢＬ番号指定で検索
			'''データ取得
			StrSQL = "SELECT MAX(RegisterDate) RegisterDate, NULL ContNo, BLNo, MAX(LatestSentTime) LatestSentTime "
			StrSQL = StrSQL & " FROM TargetContainers "
			StrSQL = StrSQL & " WHERE UserCode='"& USER &"' AND Process='R' "
			StrSQL = StrSQL & " AND BLNo is not NULL AND BLNo like '%" & SortKey & "'"
			StrSQL = StrSQL & " Group BY BLNo"
			StrSQL = StrSQL & " ORDER BY RegisterDate DESC"

			ObjRS.Open StrSQL, ObjConn, 3, 1
			if err <> 0 then
				DisConnDBH ObjConn, ObjRS
				jumpErrorP "1","c101","01","ステータス配信依頼中一覧","101","SQL:<BR>"&strSQL
			end if
			'''対象件数設定
			Num = ObjRS.RecordCount

	End Select

	if Num > 0 then
		ReDim DtTbl(Num+1)
		DtTbl(0)=Array("No.","Register Date","Container No.","BL No.","Last Delivery Date")

		i = 1
		Do Until ObjRS.EOF
			DtTbl(i)=Array("","","","","")
			DtTbl(i)(0) = i
			DtTbl(i)(1) = Left(ObjRS("RegisterDate"),10)
			if IsNull(ObjRS("ContNo")) then
				DtTbl(i)(2) = "　"
				DtTbl(i)(3) = Trim(ObjRS("BLNo"))
			else
				DtTbl(i)(2) = Trim(ObjRS("ContNo"))
				DtTbl(i)(3) = "　"
			end if
			if Trim(ObjRS("LatestSentTime")) <> "" then
				DtTbl(i)(4) = Left(ObjRS("LatestSentTime"),19)
			else
				DtTbl(i)(4) = "　"
			end if
			i=i+1
			ObjRS.MoveNext
		Loop
	end if

	ObjRS.close

	'''DB接続解除
	DisConnDBH ObjConn, ObjRS
	'''エラートラップ解除
	on error goto 0

	'''ログ出力
	WriteLogH "c101", "ステータス配信依頼中一覧", "01",""

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>Import Status Delivery Request</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//データが無い場合の表示制御
function view(){
}
//削除ほか
function GoDelEtc(contBLNo,contORBL){
	f=document.sst100F;
	f.ContBLNo.value=contBLNo;
	f.ContORBL.value=contORBL;
	f.action="sst220.asp";
	newWin = window.open("", "ReEntry", "status=yes,width=450,height=180,resizable=yes");
	f.target="ReEntry";
	f.submit();
	f.target="_self";
}
//コンテナ情報照会
function GoConinf(contBLNo,contORBL){
	f=document.sst100F;
	f.ContBLNo.value=contBLNo;
	f.ContORBL.value=contORBL;
	f.action="sst900.asp";
	newWin = window.open("", "ConInfo", "status=yes,scrollbars=yes,resizable=yes,menubar=yes");
	f.target="ConInfo";
	f.submit();
	f.target="_self";
}
//検索
function SearchC(SortFlag,Key){
	f=document.sst100F;
	f.SortFlag.value=SortFlag;
	f.SortKey.value=Key;
	f.target="_self";
	f.action="./sst100L.asp";
	f.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" onLoad="view()">
<!-------------ステータス配信依頼中画面List--------------------------->
<TABLE border="1" cellPadding="3" cellSpacing="0" width="500">
<% if Num > 0 Then %>
	<TR class=bga>
		<TH><%=DtTbl(0)(0)%></TH>
		<TH><%=DtTbl(0)(1)%></TH>
		<TH><%=DtTbl(0)(2)%></TH>
		<TH><%=DtTbl(0)(3)%></TH>
		<TH><%=DtTbl(0)(4)%></TH>
	</TR>
	<% for j=1 to Num %>
	<TR class=bgw>
		<% if DtTbl(j)(3) = "　" then %>		<%' ＢＬ番号欄表示用に全角空白がセットされている場合。すなわちコンテナ番号を表示する行の場合 %>
		<TD align="center"><A HREF="JavaScript:GoDelEtc('<%=DtTbl(j)(2)%>',1);"><%=DtTbl(j)(0)%></A></TD>
		<TD><%=DtTbl(j)(1)%></TD>
		<TD><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(2)%>',1);"><%=DtTbl(j)(2)%></A></TD>
		<TD><%=DtTbl(j)(3)%></TD>
		<TD><%=DtTbl(j)(4)%></TD>
		<% else %>
		<TD align="center"><A HREF="JavaScript:GoDelEtc('<%=DtTbl(j)(3)%>',2);"><%=DtTbl(j)(0)%></A></TD>
		<TD><%=DtTbl(j)(1)%></TD>
		<TD><%=DtTbl(j)(2)%></TD>
		<TD><A HREF="JavaScript:GoConinf('<%=DtTbl(j)(3)%>',2);"><%=DtTbl(j)(3)%></A></TD>
		<TD><%=DtTbl(j)(4)%></TD>
		<% end if %>
	</TR>
	<% next %>
<% else %>
  <TR class=bgw><TD>No email delivery request currently</TD></TR>
<% end if %>
</TABLE>
<Form name="sst100F" method="POST">
	<INPUT type="hidden" name="ContBLNo" value="" >
	<INPUT type="hidden" name="ContORBL" value="" >
	<INPUT type="hidden" name="SortFlag" value="<%=SortFlag%>" >
	<INPUT type="hidden" name="SortKey" value="<%=SortKey%>" >
</Form>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
