<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	【海貨入力】	データ一覧表示
	'
%>

<%
	' 検索一覧表示最大値
	Dim sUser,sUserNo
    sUser   = UCase(Trim(Request.form("user")))
    sUserNo = UCase(Trim(Request.form("userno")))

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "nyuryoku-kaika.asp"             'メニュー画面へ
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' 詳細表示行のデータの取得
    Dim strData()
    LineNo=0
    Do While Not ti.AtEndOfStream
        strTemp=ti.ReadLine
        ReDim Preserve strData(LineNo)
        strData(LineNo) = strTemp
        LineNo=LineNo+1
    Loop
    ti.Close
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function formSend(formname){
		window.document.forms[formname].submit();
	}

<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここからログイン入力画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika4t.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
					&nbsp;&nbsp;<%=GetUpdateTime(fs)%>
					</font>
				</td>
			</tr>
		</table>

      <table>
        <tr>
          <td> 

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>更新対象一覧</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
            <br>
<% 
	If LineNo=0 Then

	    ' エラーメッセージの表示
		Dim strErrMsg
		strErrMsg = "削除完了しました。<BR>表示出来るデータが存在しません。"
	    DispInformationMessage strErrMsg 
%>

          </td>
        </tr>
      </table>
    <br>
	<br>

<%
	Else
%>
<table border="0" cellspacing=0 cellpadding=0><tr><td>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※1) クリックで貨物情報を変更</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※2）96=HC</font></td>
          </tr>
        </table>

</td></tr><tr><td>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>船名</td>
                <td nowrap>Voyage No.</td>
                <td nowrap>荷主コード</td>
                <td nowrap>荷主管理<BR>番号<font size="-1"><sup>(※1)</sup></font></td>
                <td nowrap>Booking No.</td>
                <td nowrap>指定陸運<BR>業者コード</td>
                <td nowrap>空コン倉庫<BR>到着指定日時</td>
                <td nowrap>倉庫略称</td>
                <td nowrap>ＣＹ搬入<BR>指定日</td>
                <td nowrap>サイズ</td>
                <td nowrap>タイプ</td>
                <td nowrap>高さ<BR><font size="-1"><sup>(※2)</sup></font></td>
                <td nowrap>空コン<BR>ピック場所</td>
              </tr>

<%

	For i = 1 to LineNo
	    'トランザクションファイル作成
	    anyTmp=Split(strData(i-1),",")
%>
              <tr bgcolor="#FFFFFF"> 
				<td nowrap align=center valign=middle><%=anyTmp(0)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(1)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(2)%></td>

			<form method=post action="ms-kaika-expinfo-new.asp?kind=0">
				<td nowrap align=center valign=middle>
					<input type=hidden name="vslcode"	 value="<%=anyTmp(0)%>">
					<input type=hidden name="voyctrl"	 value="<%=anyTmp(1)%>">
					<input type=hidden name="user"		 value="<%=anyTmp(2)%>">
					<input type=hidden name="userno"	 value="<%=anyTmp(3)%>">
					<input type=hidden name="booking"	 value="<%=anyTmp(4)%>">
					<input type=hidden name="tradercode" value="<%=anyTmp(5)%>">
					<input type=hidden name="arvtime"	 value="<%=anyTmp(6)%>">
					<input type=hidden name="cydate"	 value="<%=anyTmp(7)%>">
					<input type=hidden name="size"		 value="<%=anyTmp(8)%>">
					<input type=hidden name="type"		 value="<%=anyTmp(9)%>">
					<input type=hidden name="height"	 value="<%=anyTmp(10)%>">
					<input type=hidden name="remark"	 value="<%=anyTmp(11)%>">
					<input type=hidden name="pickplace"	 value="<%=anyTmp(12)%>">
					<input type=hidden name="lineno"	 value="<%=i%>">

					<a href="JavaScript:formSend(<%=i%>)"><%=anyTmp(3)%></a>
				</td>
			</form>

<%
		For j = 0 to 12
			If anyTmp(j)="" Then
				anyTmp(j) = "<BR>"
			End If
		Next
%>
				<td nowrap align=center valign=middle><%=anyTmp(4)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(5)%></td>
<%
		If Not anyTmp(6)="<BR>" Then
			anyTmp(6) = Right(anyTmp(6),11)
		End If
		If Not anyTmp(7)="<BR>" Then
			anyTmp(7) = Right(anyTmp(7),5)
		End If
%>
				<td nowrap align=right  valign=middle><%=anyTmp(6)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(11)%></td>
				<td nowrap align=right  valign=middle><%=anyTmp(7)%></td>

				<td nowrap align=right  valign=middle><%=anyTmp(8)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(9)%></td>
				<td nowrap align=right  valign=middle><%=anyTmp(10)%></td>
				<td nowrap align=center valign=middle><%=anyTmp(12)%></td>
			  </tr>
<%
	Next
%>
			</table>
</td></tr></table>
		  </td>
		</tr>
	  </table>

    <br>

<% End If %>
		</td>
	</tr>
	<tr>
		<td valign="bottom"> 
<%
    DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------ログイン画面終わり--------------------------->
<%
    DispMenuBarBack "ms-kaika-expinfo.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log作成
    WriteLog fs, "4104","海貨入力輸出貨物情報-更新対象一覧", "00", ","
%>
