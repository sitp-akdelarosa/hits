<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "IMPORT", "impentry.asp"

    ' 表示モードの取得
    Dim bDispMode          ' true=コンテナ検索 / false=BL検索
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
    Else
        bDispMode = false
    End If
'☆☆☆ Add by nics 2010.02.17
    Dim USER
	USER    = Session.Contents("userid")
'☆☆☆ end of Add by nics 2010.02.17

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
    strFileName="../temp/" & strFileName

    ' 輸入コンテナ照会リスト表示
    WriteLog fs, "2005","輸入コンテナ照会-搬出後の位置情報＆基本情報","00", ","

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
          <td rowspan=2><img src="gif/implistt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
'	DisplayCodeListButton
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
                <td nowrap><b>Location information and Basis information after CY out.</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
            <br>
        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
<!-- mod by nics 2009.03.05 -->
<!--            <td>(*1)Display datails when clicking a container No. </td>-->
            <td>(*1)Display details when clicking a container No. </td>
<!-- end of mod by nics 2009.03.05 -->
            <td width="15"><BR></td>
            <td> (*2) 96=HC</td>
            <td width="15"><BR></td>
            <td nowrap> (*3)Presence of dangerous articles related to Fire Defense Law</td>
          </tr>
        </table>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap rowspan='3'>BL "
        Response.Write "No.</td>"
    End If
%>
                <td nowrap rowspan="3">Container No.<font size="-1"><sup>(*1)</sup></font></td>
<!-- Mod-S MES Aoyagi 2010.11.23 -->
<!--                <td nowrap colspan="5">Basis information</td> -->
		<td nowrap colspan="6">Basis information</td>
<!-- Mod-E MES Aoyagi 2010.11.23 -->
<!-- mod by nics 2009.03.05 -->
<!--                <td nowrap colspan="2">Terminal</td>-->
                <td nowrap colspan="3">Terminal</td>
<!-- end of mod by nics 2009.03.05 -->
                <td nowrap colspan="2">Stock Yard</td>
<!-- mod by nics 2009.03.05 -->
<!--                <td nowrap colspan="5">Conveyance by land</td>-->
<!-- mod by nics 2010.02.17 -->
<!--                <td nowrap colspan="4">Conveyance by land</td>-->
<%
    If Not bDispMode  or USER <> "" Then
        Response.Write "<td nowrap colspan='4'>"
        Response.Write "Conveyance by land</td>"
	Else
        Response.Write "<td nowrap colspan='3'>"
        Response.Write "Conveyance by land</td>"
    End If
%>
<!-- mod by nics 2010.02.17 -->
<!-- end of mod by nics 2009.03.05 -->
              </tr>
              <tr bgcolor="#FFCC33" align="center"> 
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">size</td>
<!-- Add-S MES Aoyagi 2010.11.27 コンテナタイプ追加 -->
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">type</td>
<!-- Add-E MES Aoyagi 2010.11.27 コンテナタイプ追加 -->
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">height(<BR>
                  <font size="-1"><sup>(*2)</sup></font></td>
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">Reefer</td>
                <td nowrap bgcolor="#FFFFCC" align="center" rowspan="2">GW(t)</td>
                <td nowrap align="center" bgcolor="#FFFFCC" rowspan="2">Hazard<BR><font size="-1"><sup>(*3)</sup></font></td>
                <td nowrap bgcolor="#FFFFCC" rowspan="2">Delivery <br>
                  permission</td>
<!-- add by nics 2009.03.05 -->
                <td nowrap rowspan="2" bgcolor="#FFFFCC"><font color="#000000">Delivery Yard<br>(code)</font></td>
                <td nowrap rowspan="2" bgcolor="#FFFFCC"><font color="#000000">Operater</font></td>
<!-- end of add by nics 2009.03.05 -->
<!-- commented by nics 2009.03.05
                <td nowrap bgcolor="#FFFFCC" rowspan="2">Delivery <br>
                  yard</td>
end of comment by nics 2009.03.05 -->
                <td nowrap bgcolor="#FFFFCC" colspan="2">SY out</td>
<!-- commented by nics 2009.03.05
                <td nowrap bgcolor="#FFFFCC" colspan="2">Warehouse arrival </td>
                <td nowrap bgcolor="#FFFFCC" rowspan="2">DeVanning<br>
                  time</td>
end of comment by nics 2009.03.05 -->
                <td nowrap bgcolor="#FFFFCC" rowspan="2">Empty container<br>
                  return</td>
                <td nowrap bgcolor="#FFFFCC" rowspan="2">Return place</td>
<!-- add by nics 2009.03.05 -->
<!-- mod by nics 2010.02.17 -->
<!--                <td nowrap bgcolor="#FFFFCC" rowspan="2">Detention<br>Free Time</td>	-->
<%
    If Not bDispMode  or USER <> "" Then
        Response.Write "<td nowrap bgcolor='#FFFFCC' rowspan='2'>"
        Response.Write "Detention<br>Free Time</td>"
    End If
%>
<!-- end of mod by nics 2010.02.17 -->
<!-- end of add by nics 2009.03.05 -->
              </tr>
              <tr bgcolor="#FFFFCC" align="center"> 
                <td nowrap>Reservation</td>
                <td nowrap>Actual</td>
<!-- commented by nics 2009.03.05
                <td nowrap>Instruction</td>
                <td nowrap>Actual</td>
end of comment by nics 2009.03.05 -->
              </tr>
<!-- ここからデータ繰り返し -->
<% ' 表示ファイルのレコードがある間繰り返す
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF">
<% ' BL No
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
    Response.Write "<a href='impdetail.asp?line=" & LineNo & "&return=2'>" & anyTmp(1) & "</a>"
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' サイズ
    If anyTmp(23)<>"" Then
        Response.Write anyTmp(23)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- Add-S MES Aoyagi 2010.11.23 -->
<% ' タイプ
    If anyTmp(46)<>"" Then
        Response.Write anyTmp(46)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- Add-E MES Aoyagi 2010.11.23 -->

<% ' 高さ
    If anyTmp(24)<>"" Then
        Response.Write anyTmp(24)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' リーファー
    If anyTmp(25)="R" Then
        Response.Write "○"
    ElseIf anyTmp(25)<>"" Then
        Response.Write "−"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 重量
    If anyTmp(26)<>"" And anyTmp(26)<>"0" Then
        dWeight=anyTmp(26) / 10
        Response.Write dWeight
    Else
        Response.Write "−"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 危険物
    If anyTmp(27)="H" Then
        Response.Write "○"
    ElseIf anyTmp(27)<>"" Then
        Response.Write "−"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ターミナル搬出可否
    If anyTmp(4)="Y" Then
        Response.Write "Permitted"
    ElseIf anyTmp(4)="S" Then
        Response.Write "Delivered"
    Else
        Response.Write "Stopped"
    End If
%>
                </td>
<!-- commented by nics 2009.03.05
                <td nowrap align=center valign=middle>
<% ' ターミナル - 搬出場所
    If anyTmp(5)<>"" Then
        Response.Write anyTmp(5)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
end of comment by nics 2009.03.05 -->
<!-- add by nics 2009.03.05 -->
                     <td nowrap align=center valign=middle>
<% ' 搬出ターミナル(蔵置場所コード)
    strDisp = "<br>"
    If anyTmp(5) <> "" Then
        strDisp = anyTmp(5)
        If anyTmp(43) <> "" Then
            strDisp = strDisp & "<br>(" & anyTmp(43) & ")"
        End If
    End If
    Response.Write strDisp
%>
                     </td>
                     <td nowrap align=center valign=middle>
<% ' 担当オペレータ
    If anyTmp(45)<>"" Then
        Response.Write anyTmp(45)
    Else
        Response.Write "<br>"
    End If
%>
                     </td>
<!-- end of add by nics 2009.03.05 -->
                <td align="center" nowrap>
<% ' ストックヤード - 搬出予約 $追加
    sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
    If Left(sTemp,1)="<" Then
        Response.Write sTemp
    Else
        Response.Write sColor
        Response.Write Left(sTemp,5) & "<br>" & Mid(sTemp,7)
        If sColor<>"" Then
            Response.Write "</font>"
        End If
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ストックヤード - 搬出完了
    Response.Write DispDateTimeCell(anyTmp(30),10)
%>
                </td>
<!-- commented by nics 2009.03.05
                <td align="center" nowrap>
<% ' 陸上運送 - 倉庫到着スケジュール
    If anyTmp(34)<>"" Then
        If anyTmp(34)<anyTmp(14) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(34),10)
    If anyTmp(34)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上輸送 - 倉庫到着完了
    Response.Write DispDateTimeCell(anyTmp(14),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 陸上輸送 - デバン完了
    Response.Write DispDateTimeCell(anyTmp(15),10)
%>
                </td>
end of comment by nics 2009.03.05 -->
                <td nowrap align=center valign=middle>
<% ' 陸上輸送 - 空コン返却完了
    Response.Write DispDateTimeCell(anyTmp(16),10)
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' 返却場所
    If anyTmp(10)<>"" Then
        Response.Write anyTmp(10)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add by nics 2009.03.05 -->
<!-- mod by nics 2010.02.17 -->
<!--                <td align="center" nowrap>	-->
<% ' ディテンションフリータイム
    ' anyTmp(39) ← ディテンションフリータイム
    ' anyTmp(16) ← 空バン返却日時[yyyy/mm/dd hh:nn]
    ' anyTmp(44) ← 空バン返却予定日[yyyy/mm/dd]
	If Not bDispMode  or USER <> "" Then
	    Response.Write "<td align='center' nowrap>"
	    strDisp = anyTmp(39)
	    strColor = "#000000"    ' 黒
	    ' 空バン返却日時が設定されている場合
	    If anyTmp(16) <> "" Then
	        ' 空バン返却日時＜システム日付の場合
	        If Left(anyTmp(16),10) < DispDateTime(Now,10) Then
	            strDisp = "−"
	        End If
	    ' 空バン返却日時が設定されていない場合
	    Else
	        ' 空バン返却予定日時が設定されている場合
	        If IsDate(anyTmp(44)) Then
	            ' 空バン返却予定日≦システム日付の場合
	            If anyTmp(44) <= DispDateTime(Now,10) Then
	                strColor = "#FF0000"    ' 赤
	            ' (空バン返却予定日−2日)≦システム日付の場合
	            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(44))),10) <= DispDateTime(Now,10) Then
	                strColor = "#FFA500"    ' 黄
	            End If
	        End If
	    End If
	    Response.Write "<font color='" & strColor & "'>"
	    Response.Write strDisp
	    Response.Write "</font>"
	'add by nics 2010.02.17
	    Response.Write "　</td>"
	End If
	'end of mod by nics 2010.02.17
%>
<!-- del by nics 2010.02.17 -->
<!--              　</td>	-->
<!-- del by nics 2010.02.17 -->
<!-- end of add by nics 2009.03.05 -->
              </tr>
<%
    Loop
%>
<!-- ここまで -->
            </table>
<form>
      <input type=button value='Display Update' OnClick="JavaScript:window.location.href='impreload.asp?request=implist2.asp'">
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
    DispMenuBarBack "implist.asp"
%>
</body>
</html>
