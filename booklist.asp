<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "EXPORT", "expentry.asp"

	Dim strBookingNo
	strBookingNo = ""
'2006/03/06 add-s h.matsuda
  dim ShipLine,ShoriMode
  ShoriMode = ""
  ShipLine = ""
'2006/03/06 add-e h.matsuda

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
    WriteLog fs, "1010","ブッキング情報照会-ブッキング情報一覧","00", ","

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)
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
          <td rowspan=2><img src="gif/bookingt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.17
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.17
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.17 -->
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
                <td nowrap><b>ブッキング情報一覧</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>

            <table border="0" cellspacing="2" cellpadding="1">
              <tr> 
                <td width="15"><BR></td>
                <td><font color="#000000" size="-1">（※1）96=HC</font></td>
                <td width="15"><BR></td>
                <td><font color="#000000" size="-1">（※2) クリックでピックアップ済コンテナNo.を表示</font></td>
              </tr>
            </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 

                <td nowrap>Booking No.</td>
                <td nowrap>船社</td>
                <td nowrap>船名</td>
				<td nowrap>Voyage No.</td>
                <td nowrap>仕向港</td>
				<td nowrap>空コン搬出場所</td>
<!-- 2008.01.12 NICS START -->
				<td nowrap>CYカット</td>
<!-- 2008.01.12 NICS END -->
				<td nowrap>サイズ</td>
				<td nowrap>タイプ</td>
				<td nowrap>高さ<font size="-1"><sup>(※1)</sup></font></td>
<!-- I20040223 S -->
				<td nowrap>材質</td>
<!-- I20040223 E -->
				<td nowrap>予約<BR>本数</td>
				<td nowrap>ピックアップ済<BR>本数<font size="-1"><sup>(※2)</sup></font></td>
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
<% ' Booking No
        If strBooking<>anyTmp(1) Then
            Response.Write anyTmp(1)
            strBooking=anyTmp(1)

			'Reload用
			If strBookingNo="" Then
				strBookingNo = anyTmp(1)
				'2006/03/06 add-s h.matsuda ブッキング重複処理に対応
				  if ubound(anyTmp)>14 then
					if trim(anyTmp(15))="ShoriMode=EMoutInf" then
						ShipLine = trim(anyTmp(14))
						ShoriMode = trim(mid(anyTmp(15),11))
					end if
				  end if
				'2006/03/06 add-e h.matsuda
			Else
				strBookingNo = strBookingNo & "," & anyTmp(1)
			End If
        Else
            Response.Write "<br>"
        End If
%>
				</td>
                <td nowrap align=center valign=middle>
<% ' 船社
        If anyTmp(2)<>"" Then
		    Response.Write anyTmp(2)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 船名
        If anyTmp(3)<>"" Then
		    Response.Write anyTmp(3)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' Voyage No.
        If anyTmp(4)<>"" Then
		    Response.Write anyTmp(4)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 仕向港 - 着岸予定
        If anyTmp(5)<>"" Then
		    Response.Write anyTmp(5)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 空コン搬出場所
        If anyTmp(6)<>"" Then
		    Response.Write anyTmp(6)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- 2008.01.12 NICS START -->
<% ' CYカット
        If anyTmp(14)<>"" Then
		    Response.Write anyTmp(14)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- 2008.01.12 NICS END -->

<% ' サイズ
        If anyTmp(7)<>"" Then
		    Response.Write anyTmp(7)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' タイプ
        If anyTmp(8)<>"" Then
		    Response.Write anyTmp(8)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 高さ
        If anyTmp(9)<>"" Then
		    Response.Write anyTmp(9)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- I20040223 S -->
<% ' 材質
        If anyTmp(12)<>"" Then
		    Response.Write anyTmp(12)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- I20040223 E -->

<% ' 予約本数
        If anyTmp(10)<>"" Then
		    Response.Write anyTmp(10)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 搬出済本数
        If anyTmp(11)<>"0" Then
		    Response.Write "<a href='#' onClick='JavaScript:window.open(""bookpick.asp?line=" & LineNo &_
						   """,""pickcont"",""scrollbars=yes,resizable=yes,width=500,height=380"")'>" &_
						   anyTmp(11) & "</a>"
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
<form method=post action="bookcheck.asp">
	  <input type=hidden name="booking" value="<%=strBookingNo%>">
<% 'Mod-s 2006/03/06 h.matsuda%>
	  <INPUT type=hidden name="ShipLine" value="<%=ShipLine%>">
	  <INPUT type=hidden name="ShoriMode" value="<%=ShoriMode%>">
<%'Mod-e 2006/03/06 h.matsuda%>
      <input type=submit value="表示データの更新">
</form>
          </td>
        </tr>
      </table>
      <form action="bookcsvout.asp"><input type="submit" value="CSVファイル出力">
    　<a href="help23.asp">CSVファイル出力とは？</a> 
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
    DispMenuBarBack "bookentry.asp"
%>
</body>
</html>

