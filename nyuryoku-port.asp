<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-in1.asp"

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 表示ファイルの取得
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' 引数指定のないとき
        strFileName="test.csv"
    End If
    strFileName="./temp/" & strFileName

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' 本船動静基本情報の取得
    If Not ti.AtEndOfStream Then
        anyTmp=Split(ti.ReadLine,",")
    End If

    ' 詳細表示行のデータ数の取得
    If Not ti.AtEndOfStream Then
        iCount=CInt(ti.ReadLine)
    End If
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<meta http-equiv="Pragma" content="no-cache">
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
          <td rowspan=2><img src="gif/nyuryoku-s.gif" width="506" height="73"></td>
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
<!-- commented by seiko-denki 2003.07.18--->
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
<!-- End of Addition by seiko-denki 2003.07.18--->
		<BR>
		<BR>
		<BR>
<table border=0><tr><td align=left>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>本船動静一覧</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr>
          <td align=left>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
                <td bgcolor="#000099" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>船社</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' 船社名の表示
    Response.Write anyTmp(1)
%>
                </td>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>船名</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' 船名の表示
    Response.Write anyTmp(3)
%>
                </td>


                <td bgcolor="#000099" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>Voyage No.</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' 次航の表示
    If anyTmp(5)=anyTmp(6) Then
        Response.Write anyTmp(5)
    Else
        Response.Write anyTmp(5) & "/" & anyTmp(6)
    End If
%>
                </td>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>コールサイン</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' コールサインの表示
    Response.Write anyTmp(2)
%>
                </td>
              </tr>
            </table>
			<BR>

			<table border=0 cellpadding=1><tr><td width=30></td>
			<td nowrap>
			データを更新する場合は対象となる港名を選択して下さい。<BR>
			新規ポートを追加する場合は新規ポートを選択して下さい。
			</td></tr></table>

            <table>
              <tr>
                <td>
                  <table border="1" cellspacing="1" cellpadding="3">
                    <tr bgcolor="#FFCC33">
                      <td nowrap align=center valign=middle><br></td>
                      <td nowrap align=center valign=middle>港名</td>
                      <td nowrap align=center valign=middle>着岸予定時刻</font></td>
                      <td nowrap align=center valign=middle>着岸完了時刻</font></td>
                      <td nowrap align=center valign=middle>離岸完了時刻</font></td>
                      <td nowrap align=center valign=middle>着岸 Long Schedule</font></td>
                      <td nowrap align=center valign=middle>離岸 Long Schedule</font></td>
                    </tr>
<!-- ここからデータ繰り返し -->
<%
    LineNo=1
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        Response.Write "<tr bgcolor='#FFFFFF'>"
        Response.Write "<td align=center>" & LineNo & "</td>"
        Response.Write "<td nowrap align=center valign=middle>"
        Response.Write "<a href='nyuryoku-sch.asp?line=" & LineNo & "'>"
        Response.Write anyTmp(1) & "</a></td>"
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(2)<>"" Then
            Response.Write anyTmp(2) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(3)<>"" Then
            Response.Write anyTmp(3) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(5)<>"" Then
            Response.Write anyTmp(5) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(6)<>"" Then
            Response.Write Left(anyTmp(6),10) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(7)<>"" Then
            Response.Write Left(anyTmp(7),10) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "</tr>"
        LineNo=LineNo+1
    Loop
    ti.Close
%>
<!-- ここまで -->
                    <tr bgcolor="#FFFFFF"> 
                      <td><br></td>
                      <td nowrap align=center valign=middle>
                        <a href="nyuryoku-new.asp">新規ポート</a>
                      </td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <br>
            <center>
            <form action="nyuryoku-vsl.asp"><input type="submit" name="submit" value="  送  信  ">
            </form>
            </center>
          </td>
        </tr>
      </table>
</td></tr></table>

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
    DispMenuBarBack "nyuryoku-in1.asp"
%>
</body>
</html>

<%
    ' 本船動静入力一覧
    WriteLog fs, "3003","船社／ターミナル入力-本船動静一覧","00", ","
%>
