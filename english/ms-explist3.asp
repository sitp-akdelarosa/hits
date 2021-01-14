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
    WriteLog fs, "1106","輸出コンテナ照会-荷主用情報一覧","00", ","

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '戻り画面種別を記憶
    Session.Contents("dispreturn")=3
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
          <td rowspan=2><img src="gif/expninushi.gif" width="506" height="73"></td>
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
                <td nowrap><b>輸出コンテナ情報一覧(荷主用)&nbsp;</b></td>
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
            <td><font color="#000000" size="-1">（※2）青：未読 &nbsp;&nbsp; 黒：照会済</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33">
                <td nowrap rowspan="2" valign=bottom>
<%
    If strSortKey="荷主管理番号" Then
        Response.Write "荷主管理番号<BR><img src='gif/1.gif' height=6><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "荷主管理番号<BR><img src='gif/1.gif' height=6><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist3.asp&sort=荷主管理番号'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap rowspan="2" valign=bottom>Booking No.<BR><img src="gif/1.gif" height=18></td>
                <td nowrap rowspan="2" valign=bottom>コンテナNo.<font size="-1"><sup>(※1)</sup></font><BR><img src="gif/1.gif" height=18></td>
                <td nowrap rowspan="2" valign=bottom>
<%
    If strSortKey="海貨" Then
        Response.Write "海貨<font size=-1><sup>(※2)</sup></font><BR><img src='gif/1.gif' height=6><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "海貨<font size=-1><sup>(※2)</sup></font><BR><img src='gif/1.gif' height=6><BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist3.asp&sort=海貨'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td colspan="2" nowrap>倉庫到着</td>
                <td nowrap rowspan="2">バンニング<br>完了</td>
                <td colspan="2" nowrap>CY到着</td>
                <td nowrap rowspan="2">船積<br>完了</td>
                <td nowrap rowspan="2">離岸<br>完了</td>
                <td colspan="2" nowrap>仕向港着岸</td>
              </tr>
              <tr align="center" bgcolor="#FFFF99"> 
                <td nowrap valign=bottom><font color="#000000">
<%
    If strSortKey="倉庫到着" Then
        Response.Write "指示<BR><img src='gif/sort-r.gif' vspace=2></a></font></td>"
    Else
        Response.Write "指示<BR>"
        Response.Write "<a href='ms-expreload.asp?request=ms-explist3.asp&sort=倉庫到着'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</font></td>"
    End If
%>
                <td nowrap valign=top><font color="#000000">完了</font></td>
                <td nowrap><font color="#000000">指示</font></td>
                <td nowrap><font color="#000000">完了</font></td>
                <td nowrap><font color="#000000">予定</font></td>
                <td nowrap><font color="#000000">完了</font></td>
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
<% ' 荷主情報 - 管理番号
    Response.Write anyTmp(14)
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
<% ' 海貨名
    If anyTmp(17)="" Then
        Response.Write "<font color='#0000FF'>"
    End If
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
    If anyTmp(17)="" Then
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
    Response.Write DispDateTimeCell(anyTmp(47),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' バンニング - 完了時刻
    Response.Write DispDateTimeCell(anyTmp(48),10)
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
    Response.Write DispDateTimeCell(anyTmp(49),10)
%>
                </td>
                <td align="center" nowrap> 
<% ' ターミナル - 船積完了
    Response.Write DispDateTimeCell(anyTmp(50),10)
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(51),10)
%>
                </td>
                <td align="center" nowrap>
<% ' 仕向港 - 着岸予定
    If anyTmp(53)<>"" Then
        If anyTmp(52)<>"" Then
            If anyTmp(53)<anyTmp(52) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(53),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(53),10)
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 仕向港 - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(52),10)
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- ここまで -->
            </table>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='ms-expreload.asp?request=ms-explist3.asp'">
</form>
          </td>
        </tr>
      </table>
      <form action="ms-expcsvout.asp"><input type="submit" value="CSVファイル出力">
    　<a href="help15.asp">CSVファイル出力とは？</a> 
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
