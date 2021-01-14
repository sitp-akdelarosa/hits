<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "index.asp"

    strSortKey=Session.Contents("sortkey")

	'種別
	Dim iLoginKind,sLoginKind
	iLoginKind = Request.QueryString("kind")
	Select Case iLoginKind
		Case "1"	sLoginKind = "海貨"
					iNum = "a105"
		Case "2"	sLoginKind = "陸運"
					iNum = "a106"
		Case "3"	sLoginKind = "荷主"
					iNum = "a107"
		Case "4"	sLoginKind = "港運"
					iNum = "a108"
		Case Else
	End Select

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "http://www.hits-h.com/index.asp"             'メニュー画面へ
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' 輸出コンテナ照会リスト表示
    WriteLog fs, iNum,"空コンピックアップシステム-" & sLoginKind & "用情報一覧","00", ","

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '戻り画面種別を記憶
    Session.Contents("dispreturn")=4
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから一覧画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<td rowspan=2><%
    If sLoginKind="海貨" Then
        Response.Write "<img src='gif/pickkat.gif' width='506' height='73'>"
    ElseIf sLoginKind="陸運" Then
        Response.Write "<img src='gif/pickrit.gif' width='506' height='73'>"
    ElseIf sLoginKind="荷主" Then
        Response.Write "<img src='gif/picknit.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/pickkot.gif' width='506' height='73'>"
    End If
%></td>
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
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>空コンピックアップ情報一覧(<%=sLoginKind%>用)&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
			<BR>
      <form method=post action="picklist-syori.asp">

<table border=0 cellpadding=0 cellspacing=0 width=500><tr><td>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">（※1）96=HC</font></td>
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">（※2）黒：確認済 &nbsp;&nbsp; 赤：変更 &nbsp;&nbsp; 青：未確認</font></td>

<% If sLoginKind="港運" Then %>
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">（※3）</font></td>
			<td valign=top><font color="#000000" size="-1">内容を確認したら後ろの□をクリックしてチェックマークを入れ、受取場所、または、指定日を変更したい場合は変更ボタンを、問題なければ確認ボタンを押してください。</font></td>
<% End If %>
<% If sLoginKind="陸運" Then %>
            <td width="15"><BR></td>
            <td valign=top nowrap><font color="#000000" size="-1">（※3）</font></td>
			<td valign=top><font color="#000000" size="-1">指定日を変更したい場合は、□をクリックしてチェックマークを入れ、受取指定日変更ボタンを押してください。</font></td>
<% End If %>

          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 

<% If sLoginKind<>"海貨" Then %>
                <td nowrap>関係者</td>
<% Else %>
                <td nowrap colspan=3>関係者</td>
<% End If %>
                <td nowrap colspan=3>船積情報</td>
                <td nowrap colspan=3>必要空コン</td>
                <td nowrap colspan=2>空コン受取</td>
                <td nowrap colspan=2>倉庫到着</td>
                <td nowrap colspan=2>CY搬入</td>
<% If sLoginKind="港運" Then %>
                <td nowrap rowspan=3 colspan=2>確認／変更<font size=-1><sup>（※3）</sup></font></td>
<% End If %>
<% If sLoginKind="陸運" Then %>
                <td nowrap rowspan=3>受取指定日<BR>変更<font size=-1><sup>（※3）</sup></font></td>
<% End If %>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
<% If sLoginKind<>"海貨" Then %>
                <td nowrap rowspan=2 valign=bottom>
					海貨<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="海貨" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=海貨"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
<% Else %>
                <td nowrap rowspan=2 valign=bottom>
					荷主<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="荷主" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=荷主"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
                <td nowrap rowspan=2 valign=bottom>
					港運<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="港運" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=港運"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
                <td nowrap rowspan=2 valign=bottom>
					陸運<BR>
					<img src="gif/1.gif" height=4><BR>
	<% If strSortKey="陸運" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=iLoginKind%>&sort=陸運"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
				
<% End If %>
                <td nowrap rowspan=2>船名／<BR>VoyageNo.</td>
                <td nowrap rowspan=2>BookingNo.</td>
                <td nowrap rowspan=2>荷主管理番号</td>
                <td nowrap rowspan=2>サイズ</td>
                <td nowrap rowspan=2>高さ<BR><font size=-1><sup>（※1）</sup></font></td>
                <td nowrap rowspan=2>タイプ</td>
                <td nowrap rowspan=2 valign=bottom>受取場所<font size=-1><sup>（※2）</sup></font><BR><img src="gif/1.gif" height=14></td>
                <td nowrap rowspan=2 valign=bottom>
					指定日<font size=-1><sup>（※2）</sup></font><BR>
					<img src="gif/1.gif" height=2><BR>
	<% If strSortKey="指定日" Then %>
					<img src="gif/sort-r.gif" vspace=2></td>
	<% Else %>
					<a href="pickreload.asp?kind=<%=sLoginKind%>&sort=指定日"><img src="gif/sort-b.gif" border=0 vspace=2></a></td>
	<% End If %>
                <td nowrap rowspan=2>場所</td>
                <td nowrap rowspan=2>指定日時</td>
                <td nowrap rowspan=2>場所</td>
                <td nowrap>離岸計画日</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap>カット日</td>
              </tr>

<!-- ここからデータ繰り返し -->
<% ' 表示ファイルのレコードがある間繰り返す
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF"> 

<% If sLoginKind<>"海貨" Then %>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 海貨
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<% Else %>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 荷主
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 港運
    If anyTmp(16)<>"" Then
        Response.Write anyTmp(16)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 陸運
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<% End If %>

                <td nowrap align=center valign=middle rowspan=2>
<% ' 船名／Voyage
    If anyTmp(2)<>"" Then
        Response.Write anyTmp(2) & "<BR>" & anyTmp(43)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' Booking
    If anyTmp(0)<>"" Then
        Response.Write anyTmp(0)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 荷主管理番号
    If anyTmp(14)<>"" Then
        Response.Write anyTmp(14)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' サイズ
    If anyTmp(10)<>"" Then
        Response.Write anyTmp(10)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 高さ
    If anyTmp(12)<>"" Then
        Response.Write anyTmp(12)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' タイプ
    If anyTmp(11)<>"" Then
        Response.Write anyTmp(11)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 空コン－受取場所
    If anyTmp(20)<>"" Then
		If anyTmp(26)="1" Then
	        Response.Write anyTmp(20)
		ElseIf anyTmp(27)="1" Then
	        Response.Write "<font color=""#ff0000"">" & anyTmp(20) & "</font>"
		Else
	        Response.Write "<font color=""#0000ff"">" & anyTmp(20) & "</font>"
		End If
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 空コン－指定日
    If anyTmp(24)<>"" Then
		If anyTmp(26)="1" Then
			Response.Write DispDateTimeCell(anyTmp(24),5)
		ElseIf anyTmp(28)="1" Then
	        Response.Write "<font color=""#ff0000"">" & DispDateTimeCell(anyTmp(24),5) & "</font>"
		Else
	        Response.Write "<font color=""#0000ff"">" & DispDateTimeCell(anyTmp(24),5) & "</font>"
		End If
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 倉庫－場所
    If anyTmp(13)<>"" Then
        Response.Write anyTmp(13)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' 倉庫－指定日時
    If anyTmp(15)<>"" Then
        Response.Write DispDateTimeCell(anyTmp(15),10)
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>
                <td nowrap align=center valign=middle rowspan=2>
<% ' ＣＹ－場所
    If anyTmp(22)<>"" Then
        Response.Write anyTmp(22)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ＣＹ－離岸計画日
    If anyTmp(45)<>"" Then
        Response.Write DispDateTimeCell(anyTmp(45),5)
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>

<% If sLoginKind="港運" Then %>
                <td nowrap align=center valign=middle rowspan=2>
<%		If anyTmp(26)="1" Then %>
					済</td>
<%		ElseIf anyTmp(26)="2" Then  %>
					変更</td>
<%		Else  %>
					<font color="#0000ff">未</font></td>
<%		End If  %>
<% End If %>

<% If sLoginKind="港運" Or sLoginKind="陸運" Then %>
                <td nowrap align=center valign=middle rowspan=2>
					<input type=checkbox name="check<%=LineNo%>"></td>
<% End If %>

			</tr>
			<tr bgcolor="#FFFFFF">
                <td nowrap align=center valign=middle>
<% ' ＣＹ－カット日
    If anyTmp(40)<>"" Then
        Response.Write DispDateTimeCell(anyTmp(40),5)
    Else
        Response.Write "<hr width=80% >"
    End If
%>
                </td>
			</tr>





<%
    Loop
%>
<!-- ここまで -->
            </table>

<% If sLoginKind="港運" Then %>
</td></tr><tr><td align=right>
			<input type=hidden name="allline" value="<%=LineNo%>">
			<input type=submit name="ok" value=" 確 認 "><input type=submit name="pickinput" value=" 変 更 ">
</td></tr><tr><td>
<% ElseIf sLoginKind="陸運" Then %>
</td></tr><tr><td align=right>
			<input type=hidden name="allline" value="<%=LineNo%>">
			<input type=submit name="pickinput" value=" 受取指定日変更 ">
</td></tr><tr><td>
<% Else %>
		<BR>
<% End If %>

		<input type=button value="表示データの更新" onClick="JavaScript:window.location.href='pickreload.asp?kind=<%=iLoginKind%>&sort=<%=sLoginKind%>'">

</td></tr></table>

      </form>
          </td>
        </tr>
      </table>
      </center>
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
<!-------------一覧画面終わり--------------------------->
<%
    DispMenuBarBack "pickselect.asp"
%>
</body>
</html>
