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

    ' 輸入コンテナ照会リスト表示
    WriteLog fs, "2002","輸入コンテナ照会-複数コンテナ","00", ","

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    '戻り画面種別を記憶
    Session.Contents("dispreturn")=0
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
          <td rowspan=2><img src="../gif/implistt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
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
                <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>手続き及びターミナル搬出可否情報&nbsp;</b></td>
                <td><img src="../gif/hr.gif"></td>
              </tr>
            </table>
            <br>
            <table border="0">
              <tr>
                <td>　<a href="implist1.asp">■ ターミナル搬入までの位置情報</a></td>
              </tr>
              <tr>
                <td>　<a href="implist2.asp">■ ターミナル搬出後の位置情報＆基本情報</a></td>
              </tr>
            </table>
            <table>
              <tr>
                <td>  
                  <br>

        <table border="0" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="15"><BR></td>
            <td><font color="#000000" size="-1">（※1) クリックで単独コンテナ情報を表示</font></td>
          </tr>
        </table>

                  <table border="1" cellspacing="1" cellpadding="3">
                    <tr align="center" bgcolor="#FFCC33"> 
<%
    If Not bDispMode Then
        Response.Write "<td nowrap align=center valign=middle rowspan='3' width='78'>BL No.</td>"
    End If
%>
                      <td rowspan="3" nowrap bgcolor="#FFCC33" align="center">コンテナNo.<font size="-1"><sup>(※1)</sup></font></td>
<!-- mod by nics 2009.02.24 -->
<!--                      <td colspan="5" nowrap bgcolor="#FFCC33" align="center">行政手続き</td>-->
                      <td colspan="7" nowrap bgcolor="#FFCC33" align="center">行政手続き</td>
<!-- end of mod by nics 2009.02.24 -->
                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">商取引<br>
                        DO発行</td>
                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">フリー<br>
                        タイム</td>
                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">ターミナル<br>
                        搬出可否</td>
<!-- add by nics 2009.02.24 -->
                      <td rowspan="3" nowrap bgcolor="#FFCC33"><font color="#000000">搬出ターミナル<br>(蔵置場所コード)</font></td>
                      <td rowspan="3" nowrap bgcolor="#FFCC33"><font color="#000000">本船担当<br>オペレータ</font></td>
<!-- end of add by nics 2009.02.24 -->
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
<!--                      <td rowspan="3" valign="middle" nowrap align="center" bgcolor="#FFCC33">事前入力<br>作業番号</td>-->
                    </tr>
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFFFCC" colspan="2" align="center">搬入確認時刻</td>
                      <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">動植物検疫</td>
                      <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">個別搬入</td>
                      <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">通関/<br>保税輸送</td>
<!-- add by nics 2009.02.24 -->
                      <td colspan="2" nowrap bgcolor="#FFFFCC">X線検査</td>
<!-- end of add by nics 2009.02.24 -->
                    </tr>
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFFFCC">予定</td>
                      <td nowrap bgcolor="#FFFFCC">完了</td>
<!-- add by nics 2009.02.24 -->
                      <td nowrap bgcolor="#FFFFCC">有無</td>
                      <td nowrap bgcolor="#FFFFCC">CY返却</td>
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
    Response.Write "<a href='impdetail.asp?line=" & LineNo & "'>" & anyTmp(1) & "</a>"
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' 搬入確認予定時刻
    If anyTmp(32)<>"" Then
        If anyTmp(18)<>"" Then
            If anyTmp(32)<anyTmp(18) Then
                Response.Write "<font color='#FF0000'>"
            Else
                Response.Write "<font color='#0000FF'>"
            End If
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
<% ' 搬入確認完了時刻
    Response.Write DispDateTimeCell(anyTmp(18),5)
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' 動植物
    If anyTmp(17)="S" Then
        Response.Write "×"
    ElseIf anyTmp(17)="C" Then
        Response.Write "○"
    Else
        Response.Write "－"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' 個別搬入
    If anyTmp(33)<>"" Then
        Response.Write "○"
    Else
        Response.Write "－"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' 通関／保税輸送
    If anyTmp(19)<>"" Then
        Response.Write "○"
    Else
        Response.Write "×"
    End If
%>
                      </td>
<!-- add by nics 2009.02.24 -->
                      <td nowrap align=center valign=middle>
<% ' X線有無
    If anyTmp(41)<>"" Then
        Response.Write anyTmp(41)
    Else
        Response.Write "<br>"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' X線CY返却
    If anyTmp(42)<>"" Then
        Response.Write anyTmp(42)
    Else
        Response.Write "<br>"
    End If
%>
                      </td>
<!-- end of add by nics 2009.02.24 -->
                      <td nowrap align=center valign=middle>
<% ' 商取引ＤＯ発行
    If anyTmp(21)<>"Y" Then
        Response.Write "×"
    Else
        Response.Write "○"
    End If
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' フリータイム
'☆☆☆ Mod_S  by nics 2009.02.24
'    If anyTmp(22)<>"" Then
'        If anyTmp(22)<DispDateTime(Now,10) Then
'            Response.Write "<font color='#FF0000'>"
'        Else
'            Response.Write "<font color='#000000'>"
'        End If
'        Response.Write DispDateTimeCell(anyTmp(22),5)
'        Response.Write "</font>"
'    Else
'        Response.Write DispDateTimeCell(anyTmp(22),5)
'    End If
'☆☆☆
    ' anyTmp(13) ← CY搬出日時[yyyy/mm/dd hh:nn]
    ' anyTmp(22) ← フリータイム(フリータイム延長日付)[yyyy/mm/dd]
    strDisp = DispDateTimeCell(anyTmp(22),5)
    strColor = "#000000"    ' 黒
    ' 搬出日時が設定されている場合
    If anyTmp(13) <> "" Then
        ' CY搬出日時＜システム日付の場合
        If Left(anyTmp(13),10) < DispDateTime(Now,10) Then
            strDisp = "－"
        End If
    ' 搬出日時が設定されていない場合
    Else
        ' フリータイムが設定されている場合
        If IsDate(anyTmp(22)) Then
            ' フリータイム≦システム日付の場合
            If anyTmp(22) <= DispDateTime(Now,10) Then
                strColor = "#FF0000"    ' 赤
            ' (フリータイム－２日)≦システム日付の場合
            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(22))),10) <= DispDateTime(Now,10) Then
                strColor = "#FFA500"    ' 黄
            End If
        End If
    End If
    Response.Write "<font color='" & strColor & "'>"
    Response.Write strDisp
    Response.Write "</font>"
'☆☆☆ Mod_E  by nics 2009.02.24
%>
                      </td>
                      <td nowrap align=center valign=middle>
<% ' ターミナル搬出可否
    If anyTmp(4)="Y" Then
        Response.Write "○"
    ElseIf anyTmp(4)="S" Then
        Response.Write "済"
    Else
        Response.Write "×"
    End If
%>
                      </td>
<!-- add by nics 2009.02.24 -->
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
<!-- end of add by nics 2009.02.24 -->
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26
 ' 事前入力作業番号
'   Response.Write "                      <td nowrap align=center valign=middle>"
'   Response.Write anyTmp(40)
'   Response.Write "                    　</td>"
%>
                    </tr>
<%
    Loop
%>
<!-- ここまで -->
                  </table>
<form>
      <input type=button value='表示データの更新' OnClick="JavaScript:window.location.href='impreload.asp?request=implist.asp'">
</form>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <form action="impcsvout.asp"><input type="submit" value="CSVファイル出力">
    　<a href="help06.asp">CSVファイル出力とは？</a> 
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
    DispMenuBarBack "impentry.asp"
%>
</body>
</html>
