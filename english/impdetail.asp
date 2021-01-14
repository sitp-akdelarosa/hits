<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Tempファイル属性のチェック
    CheckTempFile "IMPORT", "impentry.asp"

    ' 指定引数の取得
    Dim iLineNo
    iLineNo = CInt(Request.QueryString("line"))
    Dim iReturn
    iReturn = Session.Contents("dispreturn")

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

    ' 表示ファイルのOpen
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' 詳細表示行のデータの取得
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLineNo=LineNo Then
           Exit Do
        End If
    Loop
    bSingle = false                    '単独検索結果フラグ
    If iLineNo=1 And LineNo=1 Then
        '単独検索結果かどうかチェックする
        if ti.AtEndOfStream Then
            bSingle = true
        End If
    End If
    ti.Close

    ' 輸入コンテナ照会詳細
	WriteLog fs, "2006","輸入コンテナ照会-単独コンテナ", "00", anyTmp(1) & ","

    Session.Contents("dispcntnr")=anyTmp(1)     ' 表示コンテナNo.を記憶
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="../index.css" rel="stylesheet" type="text/css">
<SCRIPT Language="JavaScript">
<!--
function Submit(formName){
    document.forms[formName].submit();
}
// -->
<%
    DispMenuJava
%>
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function winOpen(winName,url,W,H){
  var WinD11=window.open(url,winName,'scrollbars=auto,resizable=yes,width='+W+',height='+H+'');
  WinD11.focus();
  WinD11.document.close();
}
</Script>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#0000ff" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000ff">
<!-------------ここから詳細画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/impdetailt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.17
'	DisplayCodeListButton
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
end of comment by seiko-denki 2003.07.07 -->

<!-- mod by nics 2009.03.04 -->
<!--		<table width=95% cellpadding=3>-->
		<table width=95% cellpadding=0>
<!-- end of mod by nics 2009.03.04 -->
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
<!-- mod by nics 2009.03.04 -->
<!--          <td>　<br>-->
          <td>
<!-- end of mod by nics 2009.03.04 -->
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
<% ' BL No
    If Not bDispMode Then
        Response.Write "<td bgcolor='#003399' background='gif/tableback.gif' nowrap><font color='#FFFFFF'><b>BL No</b></font></td>"
        Response.Write "<td bgcolor='#FFFFFF' nowrap>" & anyTmp(0) & "</td>"
    End If
%>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>Container No.</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<% ' コンテナNo.
    Response.Write anyTmp(1)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--			<BR>-->
			<font size="-2"><BR></font>
<!-- end of mod by nics 2009.03.04 -->
<!---------------基本情報------------------------------------------- commented by nics 2009.03.04 -->
<!---------------基本情報--------------------------------------------->
<!-- commented by nics 2009.03.04
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Basis information</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
end of comment by nics 2009.03.04 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
<!-- add by nics 2009.03.04 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">Basis<br>information</td>
<!-- end of add by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
                <td valign="top" nowrap>Item</td>
end of comment by nics 2009.03.04 -->
                <td nowrap bgcolor="#FFCC33">size</td>
<!-- mod by nics 2009.03.04 -->
<!-- Add-S MES Aoyagi 2010.11.23 -->
		<td nowrap bgcolor="#FFCC33">Type</td>
<!-- Add-E MES Aoyagi 2010.11.23 -->

<!--                <td nowrap bgcolor="#FFCC33">height<font size="-1"><sup>(*4)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">height<font size="-1"><sup>(*1)</sup></font></td>
<!-- end of mod by nics 2009.03.04 -->
                <td nowrap bgcolor="#FFCC33">Reefer</td>
                <td nowrap bgcolor="#FFCC33">GW(t)</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td valign="top" nowrap>Hazard<font size="-1"><sup>(*5)</sup></font></td>-->
                <td valign="top" nowrap>Hazard<font size="-1"><sup>(*2)</sup></font></td>
<!-- end of mod by nics 2009.03.04 -->
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap bgcolor="#FFCC33">Delivery yard</td>-->
                <td nowrap bgcolor="#FFCC33">Delivery Yard(code)</td>
<!-- end of mod by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                <td nowrap bgcolor="#FFCC33">Operater</td>
<!-- end of add by nics 2009.03.04 -->
                <td nowrap bgcolor="#FFCC33">Use Stock Yard</td>
                <td nowrap bgcolor="#FFCC33">Return place</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFFFCC" nowrap>Value</td>
end of comment by nics 2009.03.04 -->
                <td align="center" nowrap>
<% ' サイズ
    If anyTmp(23)<>"" Then
        Response.Write anyTmp(23)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>

<!-- Add-S MES Aoyagi 2010.11.23 コンテナタイプも表示 -->
<% ' タイプ
    If anyTmp(46)<>"" Then
        Response.Write anyTmp(46)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<!-- Add-E MES Aoyagi 2010.11.23 コンテナタイプも表示 -->

<% ' 高さ
    If anyTmp(24)<>"" Then
        Response.Write anyTmp(24)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' リーファー
    If anyTmp(25)="R" Then
        Response.Write "○"
    ElseIf anyTmp(25)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 総重量
    If anyTmp(26)<>"" And anyTmp(26)<>"0" Then
        dWeight=anyTmp(26) / 10
        Response.Write dWeight
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 危険物
    If anyTmp(27)="H" Then
        Response.Write "○"
    ElseIf anyTmp(27)<>"" Then
        Response.Write "－"
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td align="center" nowrap>
<% ' 搬出ターミナル
    If anyTmp(5)<>"" Then
        Response.Write anyTmp(5)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
end of comment by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                <td align="center" nowrap>
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
                <td align="center" nowrap>
<% ' 担当オペレータ
    If anyTmp(45)<>"" Then
        Response.Write anyTmp(45)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- end of add by nics 2009.03.04 -->
                <td align="center" nowrap>
<% ' ストックヤード利用 $追加
    If anyTmp(35)>="1" And anyTmp(35)<="4" Then
        Response.Write "○"
    Else
        Response.Write "×"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 返却場所
    If anyTmp(10)<>"" Then
        Response.Write anyTmp(10)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--            <table border="0" cellspacing="1" cellpadding="3">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.04 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap><font color="#000000" size="-1">(*4)96=HC</td>-->
                <td nowrap><font color="#000000" size="-1">(*1)96=HC</td>
<!-- end of mod by nics 2009.03.04 -->
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap>(*5)Presence of dangerous articles related to Fire Defense Law</td>-->
                <td nowrap><font color="#000000" size="-1">(*2)Presence of dangerous articles related to Fire Defense Law</td>
<!-- end of mod by nics 2009.03.04 -->
              </tr>
            </table>
<!-- commented by nics 2009.03.04
            <br>
end of comment by nics 2009.03.04 -->
<!---------------本船情報------------------------------------------- commented by nics 2009.03.04 -->
<!---------------本船情報--------------------------------------------->
<!-- commented by nics 2009.03.04
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Vessel Informatinon</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
end of comment by nics 2009.03.04 -->
            <table border=1 cellpadding="3" cellspacing="1">
<!-- mod by nics 2009.03.04 -->
<!--              <tr> -->
              <tr align="center" bgcolor="#FFCC33"> 
<!-- end of mod by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                <td rowspan="2" nowrap bgcolor="#6495ED">Vessel<br>information</td>
<!-- end of add by nics 2009.03.04 -->
                <td bgcolor="#FFCC33" nowrap>Ship's Line</td>
<!-- add by nics 2009.03.04 -->
                <td bgcolor="#FFCC33" nowrap>Vessel Name</td>
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
                <td bgcolor="#FFCC33" nowrap>Load Port</td>
                <td bgcolor="#FFCC33" nowrap>Previous Port</td>
              </tr>
              <tr align="center"> 
<!-- end of add by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船社
    If anyTmp(6)<>"" Then
        Response.Write anyTmp(6)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap>Vessel Name</td>
end of comment by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 船名
    If anyTmp(7)<>"" Then
        Response.Write anyTmp(7)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap>Voyage No.<font color="#FFFFFF"><b> 
                </b></font></td>
end of comment by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 次航
    If anyTmp(8)<>"" Then
        Response.Write anyTmp(8)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap>Load Port</td>
end of comment by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 仕出港
    If anyTmp(9)<>"" Then
        Response.Write anyTmp(9)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFCC33" nowrap>Previous Port</td>
end of comment by nics 2009.03.04 -->
                <td bgcolor="#FFFFFF" nowrap>
<% ' 前港
    If anyTmp(38)<>"" Then
        Response.Write anyTmp(38)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--            <br>-->
            <font size="-1"><br></font>
<!-- end of mod by nics 2009.03.04 -->
<!---------------位置情報------------------------------------------- commented by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Position information</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
end of comment by nics 2009.03.04 -->
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.03.04 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">Position<br>information</td>
<!-- end of add by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
                <td nowrap align="center" bgcolor="#FFCC33">Place</td>
end of comment by nics 2009.03.04 -->
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap bgcolor="#FFCC33">Load port<font size="-1"><sup>(*1)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">Load port<font size="-1"><sup>(*3)</sup></font></td>
<!-- end of mod by nics 2009.03.04 -->
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap bgcolor="#FFCC33">Previous port<font size="-1"><sup>(*1)</sup></font></td>-->
                <td nowrap bgcolor="#FFCC33">Previous port<font size="-1"><sup>(*4)</sup></font></td>
<!-- end of mod by nics 2009.03.04 -->
                <td colspan="4" nowrap bgcolor="#FFCC33">Terminal</td>
                <td nowrap bgcolor="#FFCC33">Stock Yard</td>
                <td colspan="3" nowrap bgcolor="#FFCC33">Conveyance by land</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.04
                <td nowrap rowspan="2" bgcolor="#FFFFCC">Process</td>
end of comment by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                <td rowspan="4" align="center"><table border="0" cellspacing="5">
                    <tr>
                      <td nowrap><a href="javascript:Submit('Form1')" class="splink">CHIWAN</a></td>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      </tr>
                    <tr>
                      <td><a href="javascript:Submit('queryForm')" class="splink">SHEKOU</a></td>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      </tr>
                    <tr>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      <td nowrap><a href="#" class="dummylink" onClick="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></td>
                      </tr>
                  </table>
                </td>
<!-- end of add by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
                <td nowrap rowspan="2" bgcolor="#FFFFCC">Departure</td>
end of comment by nics 2009.03.04 -->
                <td nowrap rowspan="2" bgcolor="#FFFFCC">Departure</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">Arrival</td>
                <td nowrap colspan="2" bgcolor="#FFFFCC">Port</td>
                <td nowrap bgcolor="#FFFFCC">SY out</td>
                <td nowrap bgcolor="#FFFFCC">Warehouse arrival time</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td nowrap rowspan="2" bgcolor="#FFFFCC">DeVanning<br>-->
<!--                  time</td>-->
                <td nowrap rowspan="2" bgcolor="#FFFFCC">Devanning<br>time</td>
<!-- end of mod by nics 2009.03.04 -->
                <td nowrap rowspan="2" bgcolor="#FFFFCC">Empty container<br>
                  return</td>
              </tr>
              <tr align="center" bgcolor="#FFFFCC">
                <td nowrap>Estimate</td>
                <td nowrap>Intended<br>
                  /Actual</td>
                <td nowrap>CY in<br>
                  time</td>
                <td nowrap>CY out<br>
                  time</td>
                <td nowrap>Reservation<br>
                  /Actual</td>
                <td nowrap>Instruction<br>
                  /Actual time</td>
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>Time</td>
                <td align="center" rowspan="2" nowrap>
<% ' 仕出港 - 離岸完了
    Response.Write DispDateTimeCell(anyTmp(11),11)
%>
                </td>
end of comment by nics 2009.03.04 -->
                <td align="center" rowspan="2" nowrap>
<% ' 前港 - 離岸完了 $追加
    Response.Write DispDateTimeCell(anyTmp(37),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル － 着岸スケジュール
    If anyTmp(31)<>"" Then
        Response.Write "<font color='#0000FF'>"
    End If
    Response.Write DispDateTimeCell(anyTmp(31),5)
    If anyTmp(31)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' ターミナル - 着岸予定
    If anyTmp(2)<>"" Then
        bLate = false
        If anyTmp(3)<>"" Then
            If anyTmp(2)<anyTmp(3) Then
                bLate = true
            End If
        End If
        If anyTmp(31)<>"" Then
            If Left(anyTmp(31),10)<Left(anyTmp(2),10) Then
                bLate = true
            End If
        End If
        If bLate Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
        Response.Write DispDateTimeCell(anyTmp(2),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(2),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル - ヤード搬入(確認)完了
    Response.Write DispDateTimeCell(anyTmp(12),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' ターミナル - ヤード搬出完了
    Response.Write DispDateTimeCell(anyTmp(13),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード - 搬出予約 $追加
    sTemp=DispReserveCell(anyTmp(35),anyTmp(36),sColor)
    Response.Write sColor
    Response.Write sTemp
    If sColor<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上運送 - 倉庫到着スケジュール
    If anyTmp(34)<>"" Then
        If anyTmp(34)<anyTmp(14) Then
            Response.Write "<font color='#FF0000'>"
        Else
            Response.Write "<font color='#0000FF'>"
        End If
    End If
    Response.Write DispDateTimeCell(anyTmp(34),11)
    If anyTmp(34)<>"" Then
        Response.Write "</font>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 陸上輸送 - デバン完了
    Response.Write DispDateTimeCell(anyTmp(15),11)
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 陸上輸送 - 空コン返却完了
    Response.Write DispDateTimeCell(anyTmp(16),11)
%>
                </td>
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' ターミナル - 着岸完了
    Response.Write DispDateTimeCell(anyTmp(3),11)
%>
                </td>
                <td align="center" nowrap>
<% ' ストックヤード - 搬出完了
    Response.Write DispDateTimeCell(anyTmp(30),11)
%>
                </td>
                <td align="center" nowrap>
<% ' 陸上輸送 - 倉庫到着完了
    Response.Write DispDateTimeCell(anyTmp(14),11)
%>
                </td>
              </tr>
            </table>
<!-- mod by nics 2009.03.04 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.04 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td>(*1)The time of Load port and the time of Previous port are local time.</td>-->
                <td><font color="#000000" size="-1">(*3)The location information in the port is displayed when clicking on a red button.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(*4)The time of Previous port is local time.</font></td>
<!-- end of mod by nics 2009.03.04 -->
              </tr>
            </table>
<!-- commented by nics 2009.03.04
            <br>
end of comment by nics 2009.03.04 -->
<!---------------手続き及び搬入確認--------------------------------- commented by nics 2009.03.04 -->
<!-----手続情報---------------->
<!-- commented by nics 2009.03.04
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>Procedure and Delivery permission information</b></td>
                <td><img src="gif/hr.gif" hspace="3"></td>
              </tr>
            </table>
            <br>
end of comment by nics 2009.03.04 -->
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="top">
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
<!-- add by nics 2009.03.04 -->
                <td rowspan="5" nowrap bgcolor="#6495ED">Procedure<br>and<br>Delivery<br>permission<br>information</td>
<!-- end of add by nics 2009.03.04 -->
<!-- commented by nics 2009.03.04
                <td rowspan="3" nowrap bgcolor="#FFCC33">Item</td>
end of comment by nics 2009.03.04 -->
<!-- mod by nics 2009.03.04 -->
<!--                <td colspan="4" nowrap bgcolor="#FFCC33">Administrative procedure</td>-->
                <td colspan="6" nowrap bgcolor="#FFCC33">Administrative procedure</td>
<!-- end of mod by nics 2009.03.04 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">DO Issue</td>
<!-- mod by nics 2010.02.17 
                <td rowspan="3" nowrap bgcolor="#FFCC33">Demurrage <br>
                  Free Time</td>-->
<%'	ＢＬ番号指定あるいは事前情報入力のみ表示する
	    If Not bDispMode  or USER <> "" Then
			Response.Write "<td rowspan='3' nowrap bgcolor='#FFCC33'>"
			Response.Write "Demurrage "
			Response.Write "<br>"
			Response.Write "Free Time"
			Response.Write "</td>"
		End If
%>
<!-- end of mod by nics 2010.02.17 -->
                <td rowspan="3" nowrap bgcolor="#FFCC33">Delivery <br>
                  permission</td>
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
<!-- mod by nics 2010.02.17
                <td rowspan="3" nowrap bgcolor="#FFCC33">Detention<br>
                  Free Time</td>-->
<%'	ＢＬ番号指定あるいは事前情報入力のみ表示する
	    If Not bDispMode  or USER <> "" Then
			Response.Write "<td rowspan='3' nowrap bgcolor='#FFCC33'>"
			Response.Write "Detention"
			Response.Write "<br>"
			Response.Write "Free Time"
			Response.Write "</td>"
		End If
%>
<!-- end of mod by nics 2010.02.17 -->
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">Confirmation time of arrival</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">Inspection</td>
                <td align="center" nowrap bgcolor="#FFFFCC" rowspan="2">Individual<br>
                  Confirmation</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">Customs clearance<br>
                  /Bond transportation<font size="-1"><sup>(*2)</sup></font></td>-->
                <td align="left" nowrap bgcolor="#FFFFCC" rowspan="2">Customs clearance<br>
                  /Bond transportation<font size="-1"><sup>(*5)</sup></font></td>
<!-- end of mod by nics 2009.03.04 -->
<!-- add by nics 2009.03.04 -->
                <td align="center" nowrap bgcolor="#FFFFCC" colspan="2">X-ray Inspection</td>
<!-- end of add by nics 2009.03.04 -->
              </tr>
              <tr> 
                <td align="center" nowrap bgcolor="#FFFFCC">Intended/Actual</td>
<!-- add by nics 2009.03.04 -->
                <td align="center" nowrap bgcolor="#FFFFCC">Need</td>
                <td align="center" nowrap bgcolor="#FFFFCC">Return</td>
<!-- end of add by nics 2009.03.04 -->
              </tr>
              <tr align="center"> 
<!-- commented by nics 2009.03.04
                <td bgcolor="#FFFFCC" rowspan="2" nowrap>Value</td>
end of comment by nics 2009.03.04 -->
                <td align="center" nowrap>
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
        Response.Write DispDateTimeCell(anyTmp(32),11)
        Response.Write "</font>"
    Else
        Response.Write DispDateTimeCell(anyTmp(32),11)
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
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
                <td align="center" rowspan="2" nowrap>
<% ' 個別搬入
    If anyTmp(33)<>"" Then
        Response.Write "○"
    Else
        Response.Write "－"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' 通関／保税輸送
    If anyTmp(19)<>"" Then
        If anyTmp(19)="O" Or anyTmp(19)="T" Then
            Response.Write "<a href='#"
            Response.Write iLineNo
            Response.Write "' onClick=""winOpen('win1','impdetail-h.asp?line="
            Response.Write iLineNo
            Response.Write "',150,150)"">○</a>"
        Else
            Response.Write "○"
        End If
    Else
        Response.Write "×"
    End If
%>
                </td>
<!-- add start by nics 2009.03.04  -->
                <td align="center" rowspan="2" nowrap>
<% ' X線有無
    If anyTmp(41)<>"" Then
        Response.Write anyTmp(41)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
                <td align="center" rowspan="2" nowrap>
<% ' X線CY返却
    If anyTmp(42)<>"" Then
        Response.Write anyTmp(42)
    Else
        Response.Write "<br>"
    End If
%>
                </td>
<!-- add end by nics 2009.03.04  -->
                <td align="center" rowspan="2" nowrap>
<% ' 商取引ＤＯ発行
    If anyTmp(21)<>"Y" Then
        Response.Write "×"
    Else
        Response.Write "○"
    End If
%>
                </td>
<!-- mod by nics 2010.02.02 
                <td align="center" rowspan="2" nowrap>-->
<% ' フリータイム ＢＬ番号指定あるいは事前情報入力のみ表示する
    If Not bDispMode  or USER <> "" Then
	    Response.Write "<td align='center' rowspan='2' nowrap>"
'☆☆☆ Mod_S  by nics 2009.03.04
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
'☆☆☆ Mod_E  by nics 2009.03.04
'Add by nics 2010.02.17 
	    Response.Write "</td>"
'end of Add by nics 2010.02.17 
	End If
'end of Mod by nics 2010.02.17 
%>
<!-- del by nics 2010.02.02 
                </td>
	 end of del by nics 2010.02.02 -->
                <td align="center" rowspan="2" nowrap>
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
<%'HiTS ver2 ADD by SEIKO n.Ooshige 2003/06/26%>
<!-- mod by nics 2010.02.17 
                <td align="center" rowspan="2" nowrap>-->
<% ' ディテンションフリータイム ＢＬ番号指定あるいは事前情報入力のみ表示する
'☆☆☆ Mod_S  by nics 2009.03.04
'    Response.Write anyTmp(39)
'☆☆☆
    ' anyTmp(39) ← ディテンションフリータイム
    ' anyTmp(16) ← 空バン返却日時[yyyy/mm/dd hh:nn]
    ' anyTmp(44) ← 空バン返却予定日[yyyy/mm/dd]
    If Not bDispMode  or USER <> "" Then
	    Response.Write "<td align='center' rowspan='2' nowrap>"
	    strDisp = anyTmp(39)
	    strColor = "#000000"    ' 黒
	    ' 空バン返却日時が設定されている場合
	    If anyTmp(16) <> "" Then
	        ' 空バン返却日時＜システム日付の場合
	        If Left(anyTmp(16),10) < DispDateTime(Now,10) Then
	            strDisp = "－"
	        End If
	    ' 空バン返却日時が設定されていない場合
	    Else
	        ' 空バン返却予定日時が設定されている場合
	        If IsDate(anyTmp(44)) Then
	            ' 空バン返却予定日≦システム日付の場合
	            If anyTmp(44) <= DispDateTime(Now,10) Then
	                strColor = "#FF0000"    ' 赤
	            ' (空バン返却予定日－2日)≦システム日付の場合
	            ElseIf DispDateTime(DateAdd("d", -2, cDate(anyTmp(44))),10) <= DispDateTime(Now,10) Then
	                strColor = "#FFA500"    ' 黄
	            End If
	        End If
	    End If
	    Response.Write "<font color='" & strColor & "'>"
	    Response.Write strDisp
	    Response.Write "</font>"
'☆☆☆ Mod_E  by nics 2009.03.04
'Add by nics 2010.02.17 
	    Response.Write "　</td>"
'end of Add by nics 2010.02.17 
	End If
'end of Mod by nics 2010.02.17 
%>
<!-- del by nics 2010.02.17 
              　</td>
	 end of del by nics 2010.02.17 -->
              </tr>
              <tr>
                <td align="center" nowrap>
<% ' 搬入確認完了時刻
    Response.Write DispDateTimeCell(anyTmp(18),5)
%>
                </td>
              </tr>
            </table>
			</td>
<!-- commented by nics 2009.03.04
                <td>&nbsp;</td>
                <td valign="top">
				<table border="1" cellpadding=" 3" cellspacing="1" bgcolor="#FFFFFF">
                  <tr>
                    <td nowrap bgcolor="#FFCC33"><center>Load Port<br>
      Information</center><font size="-1"><sup>(*3)</sup></font></td>
                  </tr>
                  <tr>
                    <td align="center"><table border="0" cellspacing="5">
                        <tr>
                          <td nowrap><a href="javascript:Submit('Form1')" class="splink">CHIWAN</a></td>
                          </tr>
                        <tr>
                          <td><a href="javascript:Submit('queryForm')" class="splink">SHEKOU</a></td>
                          </tr>

                    </table>
					</td>
                  </tr>
                </table>
				</td>
end of comment by nics 2009.03.04 -->
              </tr>
            </table>
			
			
			
<!-- mod by nics 2009.03.04 -->
<!--            <table border="0" cellspacing="2" cellpadding="1">-->
            <table border="0" cellspacing="0" cellpadding="0">
<!-- end of mod by nics 2009.03.04 -->
              <tr> 
                <td width="15">&nbsp;</td>
<!-- mod by nics 2009.03.04 -->
<!--                <td>(*2) Displayed Bond transportation period when clicking on sign '○' .<br>
                  (*3)The location information in the port is displayed when clicking on a red button.</td>-->
                <td nowrap><font color="#000000" size="-1">(*5) Displayed Bond transportation period when clicking on sign '○' .</td>
<!-- end of mod by nics 2009.03.04 -->
              </tr>
            </table>
<!-- commented by nics 2009.03.04
            <br>
end of comment by nics 2009.03.04 -->
<form>
      <input type=button value='Display Update' OnClick="JavaScript:window.location.href='impreload.asp?request=impdetail.asp'">
</form>
<form name="queryForm" method="post" target="_blank" action="http://oi.sctcn.com/Default.aspx?Action=Nav&Content=CONTAINER%20INFO.%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20&sm=CONTAINER%20INFO.">
    <input type="hidden" name="data" value="<%=anyTmp(1)%>">		
    <input type="hidden" name="OrgMenu" value="">
    <input type="hidden" name="targetPage" value="CONTAINER_INFO">
    <input type="hidden" name="nav" value="CONTAINER INFO.                         ">
</form>

<form name="Form1" method="post" action="http://www.cwcct.com/cct/conhis/con_his_infoE.aspx" id="Form1" target="_blank">
    <input type="hidden" name="Image1.x" value="0" />
    <input type="hidden" name="Image1.y" value="0" />
    <input type="hidden" name="__EVENTTARGET" value="" />
    <input type="hidden" name="__EVENTARGUMENT" value="" /> 
    <input type="hidden" name="__VIEWSTATE" value="dDwtMzMwNTk0MTUxOztsPEltYWdlMTs+Po9koK7lFKyndTfCh4n1g7KjLvsH" />
    <input type="hidden" name="cont_no" id="cont_no" value="<%=anyTmp(1)%>" />
    <input type="hidden" name="wyex" value="wyE" />
</form>

<%
    ' 検索画面から直接飛んできたときは表示する
    If bSingle Then
        Response.Write "<form action='impcsvout.asp'>"
        Response.Write "<center>"
        Response.Write "<input type='submit' name='submit' value='CSV file output'>　"
        Response.Write "</center>"
        Response.Write "</form>"
    End If
%>
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
<!-------------詳細画面終わり--------------------------->
<%
    ' 検索画面から直接飛んできたとき
    If bSingle Then
        DispMenuBarBack "impentry.asp"
    Else
        If iReturn=1 Then
            DispMenuBarBack "implist1.asp"
        ElseIf iReturn=2 Then
            DispMenuBarBack "implist2.asp"
        ElseIf iReturn=3 Then
            DispMenuBarBack "implist3.asp"
        Else
            DispMenuBarBack "implist.asp"
        End If
    End If
%>
</body>
</html>
