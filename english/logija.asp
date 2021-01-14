<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
' 集計ロジック
	On Error Resume Next

	Dim sYear,sMonth,sDay
	bKind = Request.QueryString("kind")
	If bKind=1 Then
		'' 指定した日付
		sYear=Trim(Request.form("year"))
		sMonth=Right("0" & Trim(Request.form("month")), 2)
		sDay=Right("0" & Trim(Request.form("day")), 2)
	Else
		'' 現在の日付取得
		sYear=Year(Now)
		sMonth=Right("0" & Month(Now), 2)
		sDay=Right("0" & Day(Now), 2)
	End If
	strDateTime = sYear & sMonth & sDay

	'' File System Object の生成
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	'' ログファイルの取得
	Dim strFileName
	strFileName="./ija/log/" & strDateTime & ".log"

	'' 表示ファイルのOpen
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	'' 詳細表示行のデータの取得
	Dim strLogData()
	LineNo=0
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		ReDim Preserve strLogData(LineNo)
		strLogData(LineNo) = strTemp
		LineNo=LineNo+1
	Loop
	ti.Close

	If LineNo>0 Then
		'' ログタイトル取得
		Dim iTKind
		Dim PageNum(),WkNum(),PageTitle(),SubTitle()
		Dim strTitleFileName
		strTitleFileName="./logija.txt"
		Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
		iTKind=0
		Do While Not ti.AtEndOfStream
			strTemp=ti.ReadLine
			anyTmpTitle=Split(strTemp,",")
			ReDim Preserve PageNum(iTKind)
			ReDim Preserve WkNum(iTKind)
			ReDim Preserve PageTitle(iTKind)
			ReDim Preserve SubTitle(iTKind)
			PageNum(iTKind) = anyTmpTitle(0)
			WkNum(iTKind) = anyTmpTitle(1)
			PageTitle(iTKind) = anyTmpTitle(2)
			If PageTitle(iTKind)="" Then PageTitle(iTKind)="<BR>"
			SubTitle(iTKind) = anyTmpTitle(3)
			iTKind=iTKind+1
		Loop
		ti.Close

		'' ログの集計
		ReDim Count(iTKind-1)
		For i=0 to iTKind-1
			Count(i)=0
			For j=0 to LineNo-1
				anyTmp=Split(strLogData(j),",")
				If anyTmp(1)=PageNum(i) and anyTmp(3)=WkNum(i) Then Count(i)=Count(i)+1
			Next
		Next
	End If

%>

<html>
<head>
	<title>アクセスログ集計（携帯）</title>
	<meta http-equiv="Pragma" content="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<script language="JavaScript">
	function checkFormValue(){
		if(!checkBlank(getFormValue(0))){ return showAlert("年",true); }
		if(!checkBlank(getFormValue(1))){ return showAlert("月",true); }
		if(!checkBlank(getFormValue(2))){ return showAlert("日",true); }
		if((Number(getFormValue(1))<1)||(Number(getFormValue(1))>12)) { return showAlert("月は1～12",false);}
		if((Number(getFormValue(2))<1)||(Number(getFormValue(2))>31)) { return showAlert("日は1～31",false);}
		return true;
	}
	function getFormValue(iNum){
		formvalue = window.document.input.elements[iNum].value;
		return formvalue;
	}

	function checkBlank(formvalue){
		if(formvalue == ""){ return false; }
		return true;
	}
	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "が未入力です。");
		} else {
			window.alert(strAlert + "は、どちらか一方を入力して下さい。");
		}
		return false;
	}
</script>
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<tr><td height="20"></td></tr>
<tr>
	<td valign="top">

		<center>
		<table>
		<tr> 
			<td><img src="gif/botan.gif" width="17" height="17"></td>
			<td nowrap><b>アクセスログ集計表（携帯）</b></td>
			<td><img src="gif/hr.gif" width="400" height="3"></td>
		</tr>
		</table>
		<br>
		<table border="0">
		<tr><td align="left">
			<form method="post" name="input" action="logija.asp?kind=1">
				<input type="text" name="year" maxlength="4" size="4" value=<%=Year(Now)%>>年
				<input type="text" name="month" maxlength="2" size="2">月
				<input type="text" name="day" maxlength="2" size="2">日
				<input type="submit" value="ログを表示" onClick="return checkFormValue()">
			</form>
		</td></tr>
		<tr><td align="center"><b><%=sYear & "年" & sMonth & "月" & sDay & "日"%>の情報</b></td></tr>
		<tr>
			<td align=left>
<% If LineNo>0 Then %>
				<table border="1" cellpadding="5">
					<tr>
						<th align="center" bgcolor="#6699FF">メニュー項目</th>
						<th align="center" bgcolor="#6699FF">画面</th>
						<th align="center" bgcolor="#6699FF">画面No.</th>
						<th align="center" bgcolor="#6699FF" width="100">アクセス件数</th>
					</tr>
<% For i=0 to iTKind-1 %>
					<tr>
						<td align="left"><%=PageTitle(i)%></td>
						<td align="left"><%=SubTitle(i)%></td>
						<td align="left"><%=PageNum(i)%>-<%=WkNum(i)%></td>
						<td align="right" width="85"><%=Count(i)%> </td>
					</tr>
<% Next %>
					</table>
<% Else %>
				<br><div align="center">データが1件もありません。</div><br>
<% End If %>
			</td>
		</tr>
<% If LineNo>0 Then %>
		<tr><td>
			<form action="JavaScript:window.location.reload(true)">
				<input type="hidden" name="year" value=<%=sYear%>>
				<input type="hidden" name="month" value=<%=sMonth%>>
				<input type="hidden" name="day" value=<%=sDay%>>
				<input type="submit" value="表示データの更新">
			</form>
		</td></tr>
<% End If %>
		</table>
		<a href="http://www.hits-h.com/index.asp">トップページへ戻る</a>
		<br><br>
		</center>
	</td>
</tr>
</table>
</body>
</html>