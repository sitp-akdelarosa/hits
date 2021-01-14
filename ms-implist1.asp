<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "MSIMPORT", "impentry.asp"

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "impentry.asp"             '輸入コンテナ照会トップ
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' ユーザ種類をチェックする
    strUserKind=Session.Contents("userkind")
    ' Sort条件種類をチェックする
    strSortKey=Session.Contents("sortkey")

	Dim iNum
	If strUserKind="海貨" Then
		iNum = "2104"
	ElseIf strUserKind="陸運" Then
		iNum = "2105"
	Else
		iNum = "2106"
	End If

    ' 輸入コンテナ照会リスト表示
    WriteLog fs, iNum,"輸入コンテナ照会-" & strUserKind & "用情報一覧","00", ","

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '戻り画面種別を記憶
    Session.Contents("dispreturn")=1
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

    If strUserKind="海貨" Then
        Response.Write "<img src='gif/impkaika.gif' width='506' height='73'>"
    ElseIf strUserKind="陸運" Then
        Response.Write "<img src='gif/imprikuun.gif' width='506' height='73'>"
    Else
        Response.Write "<img src='gif/impninushi.gif' width='506' height='73'>"
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
                <td nowrap><b>
<%
        Response.Write "輸入コンテナ情報一覧(" & strUserKind & "用)"
%>
                &nbsp;</b></td>
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
            <td><font color="#000000" size="-1">（※2）仕向港の時刻は、現地時間です。</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※3）青：未読 &nbsp; 黒：照会済</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap colspan="2" valign=bottom>
<%
    If strSortKey="船名" Then
        Response.Write "本船<BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "本船<BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=船名'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
<%
    If strUserKind="海貨" Then
        Response.Write "<td nowrap rowspan='3' valign=bottom>"
        If strSortKey="荷主" Then
            Response.Write "荷主<BR><img src='gif/1.gif' height=18><BR><img src='gif/sort-r.gif' vspace=2></td>"
        Else
            Response.Write "荷主<BR><img src='gif/1.gif' height=18><BR>"
            Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=荷主'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
            Response.Write "</td>"
        End If
    Else
        Response.Write "<td nowrap rowspan='3' valign=bottom>"
        If strSortKey="海貨" Then
            Response.Write "海貨<BR><img src='gif/1.gif' height=18><BR><img src='gif/sort-r.gif' vspace=2></td>"
        Else
            Response.Write "海貨<BR><img src='gif/1.gif' height=18><BR>"
            Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=海貨'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
            Response.Write "</td>"
        End If
    End If
%>
                <td nowrap rowspan="3" valign=bottom>船社<BR><img src="gif/1.gif" height=30></td>
                <td nowrap rowspan="3" valign=bottom>BL No.<BR><img src="gif/1.gif" height=30></td>
                <td nowrap rowspan="3" valign=bottom>コンテナNo.<font size="-1"><sup>(※1)</sup></font><BR><img src="gif/1.gif" height=30></td>
                <td nowrap bgcolor="#FFCC33">仕出港</td>
                <td nowrap colspan="7">ターミナル</td>
                <td nowrap colspan="3">陸上輸送</td>
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap rowspan="2" bgcolor="#FFFF99">船名</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">Voyage No.</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">離岸完了<br>
                  時刻<font size="-1"><sup>(※2)</sup></font></td>
                <td nowrap colspan="3" bgcolor="#FFFF99">着岸時刻</td>
                <td nowrap colspan="2" bgcolor="#FFFF99">搬入確認時刻</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">搬出<BR>可否</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99">搬出<br>完了</td>
                <td nowrap rowspan="2" bgcolor="#FFFF99" valign=bottom>
<%
    If strSortKey="陸運業者" Then
        Response.Write "指定陸運<br>業者<font size=-1><sup>(※3)</sup></font><BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "指定陸運<br>業者<font size=-1><sup>(※3)</sup></font><BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=陸運業者'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap colspan="2" bgcolor="#FFFF99">倉庫到着</td>
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap bgcolor="#FFFF99">計画</td>
                <td nowrap bgcolor="#FFFF99">予定</td>
                <td nowrap bgcolor="#FFFF99">完了</td>
                <td nowrap bgcolor="#FFFF99">予定</td>
                <td nowrap bgcolor="#FFFF99">完了</td>
                <td nowrap bgcolor="#FFFF99" valign=bottom>
<%
    If strSortKey="倉庫到着" Then
        Response.Write "指示<BR><img src='gif/sort-r.gif' vspace=2></td>"
    Else
        Response.Write "指示<BR>"
        Response.Write "<a href='ms-impreload.asp?request=ms-implist1.asp&sort=倉庫到着'><img src='gif/sort-b.gif' border=0 vspace=2></a>"
        Response.Write "</td>"
    End If
%>
                <td nowrap bgcolor="#FFFF99" valign=top>完了</td>
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
<% ' 船名
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' Voyage No.
    If anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<%
    If strUserKind="海貨" Then
        strTemp=anyTmp(7)
    Else
        strTemp=anyTmp(8)
    End If
    If strTemp<>"" Then
        Response.Write strTemp
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 船社
    If anyTmp(20)<>"" Then
        Response.Write anyTmp(20)
    ElseIf anyTmp(15)<>"" Then
        Response.Write anyTmp(15)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' BL No
    If anyTmp(0)<>"" Then
        Response.Write anyTmp(0)
    Else
        Response.Write "<br>"
    End If
%>
				</td>
                <td nowrap align=center valign=middle>
<% ' コンテナNo.
    Response.Write "<a href='ms-impdetail.asp?line=" & LineNo & "&return=1'>" & anyTmp(1) & "</a>"
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 仕出港 - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(41),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル － 着岸スケジュール
    If anyTmp(61)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(61),10)
    If anyTmp(61)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - 着岸予定
    If anyTmp(32)<>"" Then
        bLate = false
        If anyTmp(33)<>"" Then
            If anyTmp(32)<anyTmp(33) Then
                bLate = true
            End If
        End If
        If anyTmp(61)<>"" Then
            If anyTmp(61)<anyTmp(32) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(32),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(32),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(33),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - 搬入確認予定
    If anyTmp(62)<>"" Then
        If anyTmp(48)<>"" Then
            If anyTmp(62)<anyTmp(48) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(62),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(62),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - ヤード搬入(確認)完了
    Response.Write DispDateTimeCell(anyTmp(48),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル搬出可否
    If anyTmp(34)="Y" Then
        Response.Write "○"
    ElseIf anyTmp(34)="S" Then
        Response.Write "済"
    Else
        Response.Write "×"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - ヤード搬出完了
    Response.Write DispDateTimeCell(anyTmp(43),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上輸送 - 指定陸運業者
    If anyTmp(9)<>"" Then
        If anyTmp(14)<>"" Then
            Response.Write anyTmp(9)
        Else
            Response.Write "<font color='#0000FF'>" & anyTmp(9) & "</font>"
        End If
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上輸送 - 指示
    If anyTmp(64)<>"" Then
        strTemp=anyTmp(64)
    Else
        strTemp=anyTmp(13)
    End If
    If strTemp<>"" Then
        If strTemp<anyTmp(45) Then
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
<% ' 陸上輸送 - 完了
    Response.Write DispDateTimeCell(anyTmp(45),10)
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- ここまで -->
            </table>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='ms-impreload.asp?request=ms-implist1.asp'">
</form>
          </td>
        </tr>
      </table>
      <form action="ms-impcsvout.asp"><input type="submit" value="CSVファイル出力">
<%
    If strUserKind="海貨" Then
        Response.Write "<a href='help16.asp'>CSVファイル出力とは？</a>"
    Else
        Response.Write "<a href='help18.asp'>CSVファイル出力とは？</a>"
    End If
%>
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
    DispMenuBarBack "ms-impentry.asp"
%>
</body>
</html>
