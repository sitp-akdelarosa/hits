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
    WriteLog fs, "1005","輸出コンテナ照会-コンテナ作成に係る情報","00", ","

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
          <td rowspan=2><img src="gif/explistt.gif" width="506" height="73"></td>
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
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>コンテナ作成に係る情報　</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※1) クリックで単独コンテナ情報を表示</font></td>
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※2）96=HC</font></td>
          </tr>
        </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap rowspan='2'>Booking "
        Response.Write "No.</td>"
    End If
%>
                <td rowspan="2" bgcolor="#FFCC33" nowrap align="center">コンテナNo.<font size="-1"><sup>(※1)</sup></font></td>
<!-- mod by mes(2005/3/28) テアウェイト追加 -->
<!--                <td colspan="5" bgcolor="#FFCC33" nowrap align="center">空コンテナ</td>-->
<!--                <td colspan="6" bgcolor="#FFCC33" nowrap align="center">空コンテナ</td> -->
		<td colspan="7" bgcolor="#FFCC33" nowrap align="center">空コンテナ</td>
<!-- end mes -->
                <td colspan="6" bgcolor="#FFCC33" nowrap align="center">バンニング後コンテナ</td>
              </tr>
              <tr> 
                <td bgcolor="#FFFF99" align="center" nowrap>受取時刻</td>
                <td bgcolor="#FFFF99" align="center" nowrap>受取場所</td>
                <td bgcolor="#FFFF99" align="center" nowrap>サイズ</td>
<!-- Add-S MES Aoyagi 2010.11.23 -->
		<td bgcolor="#FFFF99" align="center" nowrap>タイプ</td>
<!-- Add-E MES Aoyagi 2010.11.23 -->
                <td bgcolor="#FFFF99" align="center" nowrap>高さ<BR><font size="-1"><sup>(※2)</sup></font></td>
<!-- add by mes(2005/3/28) テアウェイト追加 -->
                <td bgcolor="#FFFF99" align="center" nowrap>TW</td>
<!-- end mes -->
                <td bgcolor="#FFFF99" align="center" nowrap>リーファ</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>倉庫<br>
                  到着時刻</td>
end of comment by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap>シールNo.</td>
                <td bgcolor="#FFFF99" align="center" nowrap>貨物<br>重量(t)</td>
                <td bgcolor="#FFFF99" align="center" nowrap>総<br>重量(t)</td>
                <td bgcolor="#FFFF99" align="center" nowrap>バンニング<br>
                  完了時刻</td>
<!-- commented by nics 2009.02.24
                <td bgcolor="#FFFF99" align="center" nowrap>搬入<br>
                  ターミナル名</td>
end of comment by nics 2009.02.24 -->
<!-- add by nics 2009.02.24 -->
                <td bgcolor="#FFFF99" align="center" nowrap><font color="#000000">搬入ターミナル<br>(蔵置場所コード)</font></td>
                <td bgcolor="#FFFF99" align="center" nowrap><font color="#000000">本船担当<br>オペレータ</font></td>
<!-- end of add by nics 2009.02.24 -->
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
<% ' 陸上運送 - 空コン受取
    Response.Write DispDateTimeCell(anyTmp(16),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 空コン受取場所
    If anyTmp(2)<>"" Then
        Response.Write anyTmp(2)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' サイズ
    If anyTmp(3)<>"" Then
        Response.Write anyTmp(3)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- Add-S MES Aoyagi 2010.11.23 -->
<% ' タイプ
    If anyTmp(39)<>"" Then
        Response.Write anyTmp(39)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- Add-E MES Aoyagi 2010.11.23 -->

<% ' 高さ
    If anyTmp(4)<>"" Then
        Response.Write anyTmp(4)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add by mes(2005/3/28) テアウェイト追加 -->
                <td nowrap align=center valign=middle>
<% ' テアウェイト(TW)
    If anyTmp(32)<>"" And anyTmp(32)>0 Then
    	If anyTmp(32)<100 then
	        dWeight=anyTmp(32) * 100
	    Else
	        dWeight=anyTmp(32)
	    End If
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
<!-- end mes -->
                <td nowrap align=center valign=middle>
<% ' リーファー
    If anyTmp(5)="R" Then
        Response.Write "○"
    ElseIf anyTmp(5)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' 陸上運送 - 倉庫到着
    Response.Write DispDateTimeCell(anyTmp(17),10)
%>
                </td>
end of comment by nics 2009.02.24 -->
                <td nowrap align=center valign=middle>
<% ' シールNo.
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 貨物重量 $追加
    If anyTmp(27)<>"" And anyTmp(27)<>"0" Then
        dWeight=anyTmp(27) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 総重量
    If anyTmp(8)<>"" And anyTmp(8)<>"0" Then
        dWeight=anyTmp(8) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上運送 - バンニング
    Response.Write DispDateTimeCell(anyTmp(18),10)
%>
                </td>
<!-- commented by nics 2009.02.24
                <td nowrap align=center valign=middle>
<% ' 搬入ターミナル名
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
end of comment by nics 2009.02.24 -->
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
              </tr>
<%
    Loop
%>
<!-- ここまで -->
            </table>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='expreload.asp?request=explist2.asp'">
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
    DispMenuBarBack "explist.asp"
%>
</body>
</html>
