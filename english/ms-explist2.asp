<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' ソートモードの取得
    Dim strSortKey
    strSortKey=Session.Contents("sortkey")

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
    WriteLog fs, "1105","輸出コンテナ照会-陸運用情報一覧","00", ","

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '戻り画面種別を記憶
    Session.Contents("dispreturn")=2
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
          <td rowspan=2><img src="gif/exprikuun.gif" width="506" height="73"></td>
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
                <td nowrap><b>輸出コンテナ情報一覧(陸運用)&nbsp;</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>
			&nbsp;&nbsp;&nbsp;&nbsp;項目内のボタンを押すと、その項目の値でソートされます。<BR><BR>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※1) クリックで単独コンテナ情報を表示</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※2）96=HC</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※3）クリックで完了時刻入力画面へ</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap rowspan="2" valign=bottom>
<%
    If strSortKey="海貨" Then
        Response.Write "海貨<BR><img src='gif/1.gif' height=8><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "海貨<BR><img src='gif/1.gif' height=8><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist2.asp&sort=海貨'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap rowspan="2" valign=bottom>Booking No.<BR><img src="gif/1.gif" height=20></td>
                <td nowrap rowspan="2" valign=bottom>コンテナNo.<font size="-1"><sup>(※1)</sup></font><BR><img src="gif/1.gif" height=20></td>
                <td colspan="4" nowrap>空コンテナ</td>
                <td colspan="2" nowrap>倉庫到着</td>
                <td colspan="2" nowrap>バンニング後コンテナ</td>
                <td colspan="2" nowrap>CY到着</td>
                <td nowrap rowspan="2">荷主<br>名称</td>
                <td nowrap rowspan="2">倉庫略称</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap><font color="#000000">空コン<br>受取場所</font></td>
                <td nowrap><font color="#000000">サイズ</font></td>
                <td nowrap><font color="#000000">高さ<BR><font size="-1"><sup>(※2)</sup></font></font></td>
                <td nowrap><font color="#000000">リーファー</font></td>
                <td nowrap valign=bottom><font color="#000000">
<%
    If strSortKey="倉庫到着" Then
        Response.Write "指示<BR><img src='gif/sort-r.gif' vspace=2></font></td>"
    Else
        Response.Write "指示<BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist2.asp&sort=倉庫到着'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap valign=top><font color="#000000">完了<BR><font size="-1"><sup>(※3)</sup></font></font></td>
                <td nowrap><font color="#000000">完了時刻<BR><font size="-1"><sup>(※3)</sup></font></font></td>
                <td nowrap><font color="#000000">総<br>重量(t)</font></td>
                <td nowrap valign=bottom><font color="#000000">
<%
    If strSortKey="CY到着" Then
        Response.Write "指示<BR><img src='gif/sort-r.gif' vspace=2></font></td>"
    Else
        Response.Write "指示<BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist2.asp&sort=CY到着'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap valign=top><font color="#000000">完了</font></td>
              </tr>
<!-- ここからデータ繰り返し -->
<% ' 表示ファイルのレコードがある間繰り返す
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF"> 
                <td nowrap align=center valign=middle>
<% ' 海貨名
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' Booking No
    Response.Write anyTmp(0)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' コンテナNo.
    If anyTmp(1)<>"" Then
        Response.Write "<a href='ms-expdetail.asp?line=" & LineNo & "'>" & anyTmp(1) & "</a>"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 空コンテナ - 空コン受取場所
    If anyTmp(32)<>"" Then
        Response.Write anyTmp(32)
    Else
        If anyTmp(20)<>"" Then
            Response.Write "<font color='#0000FF'>" & anyTmp(20) & "</font>"
        Else
            Response.Write "<br>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 空コンテナ - サイズ
    If anyTmp(33)<>"" Then
        Response.Write anyTmp(33)
    Else
        If anyTmp(10)<>"" Then
            Response.Write "<font color='#0000FF'>" & anyTmp(10) & "</font>"
        Else
            Response.Write "<br>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 空コンテナ - 高さ
    If anyTmp(34)<>"" Then
        Response.Write anyTmp(34)
    Else
        If anyTmp(12)<>"" Then
            Response.Write "<font color='#0000FF'>" & anyTmp(12) & "</font>"
        Else
            Response.Write "<br>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 空コンテナ - リーファー
    If anyTmp(35)<>"" Then
        If anyTmp(35)="R" Then
            Response.Write "○"
        Else
            Response.Write "－"
        End If
    Else
        Response.Write "<font color='#0000FF'>"
        If anyTmp(11)<>"" Then
            If anyTmp(11)<>"RF" Then
                Response.Write "－"
            Else
                Response.Write "○"
            End If
        Else
            Response.Write "<br>"
        End If
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上運送 - 倉庫到着スケジュール
    If anyTmp(56)<>"" Then
        strTemp=anyTmp(56)
    Else
        strTemp=anyTmp(15)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(47) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(strTemp,10)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上運送 - 倉庫到着
    If anyTmp(5)<>"" Then
        Response.Write "<a href='ms-expinput.asp?kind=1&line=" & LineNo & "&request=ms-explist2.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(47),10)
    If Left(strTemp,1)="<" And anyTmp(1)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(5)<>"" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' バンニング - 完了時刻
    If anyTmp(5)<>"" Then
        Response.Write "<a href='ms-expinput.asp?kind=2&line=" & LineNo & "&request=ms-explist2.asp'>"
    End If
    strTemp = DispDateTimeCell(anyTmp(48),10)
    If Left(strTemp,1)="<" And anyTmp(1)<>"" Then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
    Else
        Response.Write strTemp
    End If
    If anyTmp(5)<>"" Then
        Response.Write "</a>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 総重量
    If anyTmp(38)<>"" And anyTmp(38)<>"0" Then
        dWeight=anyTmp(38) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上運送 - CY到着スケジュール
    If anyTmp(60)<>"" Then
        strTemp=anyTmp(60)
    Else
        strTemp=anyTmp(16)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(49) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(strTemp,5)
    If strTemp<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上運送 - CY到着
    Response.Write DispDateTimeCell(anyTmp(49),5)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 荷主名
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 倉庫略称
    If anyTmp(13)<>"" Then
        Response.Write anyTmp(13)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- ここまで -->
            </table>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='ms-expreload.asp?request=ms-explist2.asp'">
</form>
          </td>
        </tr>
      </table>
      <form action="ms-expcsvout.asp"><input type="submit" value="CSVファイル出力">
    　<a href="help14.asp">CSVファイル出力とは？</a> 
      </form>
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
    DispMenuBarBack "ms-expentry.asp"
%>
</body>
</html>
