<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "EXPORT", "expentry.asp"

    ' 表示モードの取得
    Dim bDispMode          ' true=コンテナ検索 / false=ブッキング検索
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
    Else
        bDispMode = false
    End If

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' セッションが切れているとき
        Response.Redirect "expentry.asp"             '輸出コンテナ照会トップ
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' 輸出コンテナ照会リスト表示
    WriteLog fs, "3006","仕出地情報照会-ターミナル、本船に係る情報","00", ","

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
          <td rowspan=2><img src="../gif/explistt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
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
end of comment by seiko-denki 2003.07.18 -->

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
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>ターミナル、本船に係る情報　</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
            <br>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※1) クリックで単独コンテナ情報を表示</font></td>
<!-- commented by nics 2009.02.24
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※2）仕向港の時刻は、現地時間です。</font></td>
end of comment by nics 2009.02.24 -->
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap rowspan='3'>Booking "
        Response.Write "No.</td>"
    End If
%>
                <td rowspan="3" bgcolor="#FFCC33" nowrap align="center">コンテナNo.<font size="-1"><sup>(※1)</sup></font></td>
                <td colspan="2" bgcolor="#FFCC33" nowrap align="center">搬入受付期間</td>
<!-- mod by nics 2009.02.24 -->
<!--                <td colspan="6" bgcolor="#FFCC33" nowrap align="center">ターミナル</td>-->
                <td colspan="7" bgcolor="#FFCC33" nowrap align="center">ターミナル</td>
<!-- end of mod by nics 2009.02.24 -->
                <td colspan="2" bgcolor="#FFCC33" nowrap align="center">本船</td>
<!-- commented by nics 2009.02.24
                <td colspan="2" bgcolor="#FFCC33" nowrap align="center">仕向港</td>
end of comment by nics 2009.02.24 -->
              </tr>
              <tr> 
                <td bgcolor="#FFFF99" align="center" nowrap rowspan="2">オープン日</td>
                <td bgcolor="#FFFF99" align="center" nowrap rowspan="2">クローズ日</td>
                <td bgcolor="#FFFF99" align="center" colspan="2" nowrap>搬入時刻</td>
                <td bgcolor="#FFFF99" align="center" rowspan="2" nowrap>船積<BR>
                  完了時刻</td>
<!-- mod by nics 2009.02.24 -->
<!--                <td bgcolor="#FFFF99" align="center" colspan="3" nowrap>離岸時刻</td>-->
                <td bgcolor="#FFFF99" align="center" colspan="2" nowrap>離岸時刻</td>
<!-- end of mod by nics 2009.02.24 -->
<!-- add by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center"  nowrap rowspan="2"><font color="#000000">搬入ターミナル<br>(蔵置場所コード)</font></td>
                <td bgcolor="#FFFF99" align="center"  nowrap rowspan="2"><font color="#000000">本船担当<br>オペレータ</font></td>
<!-- end of add by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap rowspan="2">船名</td>
                <td bgcolor="#FFFF99" align="center" rowspan="2" nowrap>仕出港名</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap colspan="2">着岸時刻<font size="-1"><sup>(※2)</sup></font></td>
end of comment by nics 2009.02.24 -->
              </tr>
              <tr> 
                <td bgcolor="#FFFF99" align="center" nowrap>指示</td>
                <td bgcolor="#FFFF99" align="center" nowrap>完了</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>計画</td>
end of comment by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap>予定</td>
                <td bgcolor="#FFFF99" align="center" nowrap>完了</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>予定</td>
                <td bgcolor="#FFFF99" align="center" nowrap>完了</td>
end of comment by nics 2009.02.24 -->
              </tr>
<!-- ここからデータ繰り返し -->
<% ' 表示ファイルのレコードがある間繰り返す
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF"> 
<% ' Booking No
    If Not bDispMode Then
        Response.Write "<td nowrap align=center valign=middle>"
        If strBooking<>anyTmp(0) Then
            Response.Write anyTmp(0)
            strBooking=anyTmp(0)
        Else
            Response.Write "<br>"
        End If
        Response.Write "</td>"
    End If
%>
                <td nowrap align=center valign=middle>
<% ' コンテナNo.
    Response.Write "<a href='expdetail.asp?line=" & LineNo & "'>" & anyTmp(1) & "</a>"
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' CYオープン
    Response.Write DispDateTimeCell(anyTmp(9),5)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' CYクローズ
    Response.Write DispDateTimeCell(anyTmp(10),5)
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - CY搬入指示 $追加
    If anyTmp(30)<>"" Then
        If Left(anyTmp(30),10)<Left(anyTmp(19),10) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(30),5)
    If anyTmp(30)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - CY搬入完了
    Response.Write DispDateTimeCell(anyTmp(19),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - 船積完了
    Response.Write DispDateTimeCell(anyTmp(20),10)
%>
                </td>
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' ターミナル - 離岸スケジュール
    If anyTmp(25)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(25),10)
    If anyTmp(25)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
end of comment by nics 2009.02.24 -->
                <td nowrap align=center valign=middle>
<% ' ターミナル - 離岸予定
    If anyTmp(15)<>"" Then
        bLate = false
        If anyTmp(21)<>"" Then
            If anyTmp(15)<anyTmp(21) Then
                bLate = true
            End If
        End If
        If anyTmp(25)<>"" Then
            If anyTmp(25)<anyTmp(15) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(15),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(15),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(21),10)
%>
                </td>
<!-- add by nics 2009.02.24 -->
                     <td nowrap align=center valign=middle>
<% ' 搬入ターミナル(蔵置場所コード)
    strDisp = "<br>"
    If anyTmp(6) <> "" Then
        strDisp = anyTmp(6)
        If anyTmp(36) <> "" Then
            strDisp = strDisp & "<br>(" & anyTmp(36) & ")"
        End If
    End If
    Response.Write strDisp

%>
                     </td>
                     <td nowrap align=center valign=middle>
<% ' 担当オペレータ
    If anyTmp(37)<>"" Then
        Response.Write anyTmp(37)
    Else
        Response.Write "<br>"
    End If
%>
                     </td>
<!-- end of add by nics 2009.02.24 -->
                <td nowrap align=center valign=middle>
<% ' 船名
    If anyTmp(12)<>"" Then
        Response.Write anyTmp(12)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 仕向港
    If anyTmp(14)<>"" Then
        Response.Write anyTmp(14)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' 仕向港 - 着岸予定
    If anyTmp(23)<>"" Then
        If anyTmp(22)<>"" Then
            If anyTmp(23)<anyTmp(22) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(23),10)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(23),10)
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 仕向港 - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(22),10)
%>
                </td>
end of comment by nics 2009.02.24 -->
              </tr>
<%
    Loop
%>
<!-- ここまで -->
            </table>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='expreload.asp?request=explist3.asp'">
</form>
          </td>
        </tr>
      </table>
      <br>
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
    DispMenuBarBack "explist.asp"
%>
</body>
</html>
