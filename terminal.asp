<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
''    ' セッションのチェック
''    CheckLogin "terminal.asp"

    ' ターミナル情報レコードの取得
    ConnectSvr conn, rsd

    sql = "SELECT RecWaitTime, DelWaitTime, RDWaitTime FROM Terminal WHERE Terminal='KA'"
    'SQLを発行してターミナル情報レコードを検索
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        iRecWaitTime = rsd("RecWaitTime")
        iDelWaitTime = rsd("DelWaitTime")
        iRDWaitTime = rsd("RDWaitTime")
    End If
    rsd.Close
'ADD START HiTS Ver2 By SEIKO N.Ooshige
    dim IcInTime,IcOutTime
    sql = "SELECT RecWaitTime, DelInWaitTime,DelOutWaitTime FROM Terminal2 WHERE Terminal='IC'"
    'SQLを発行してアイランドシティターミナル情報レコードを検索
    rsd.Open sql, conn, 0, 1, 1
    If Not rsd.EOF Then
        IcInTime  = rsd("RecWaitTime")					'搬入(IN受付→作業完了)
        IcOutTime = rsd("DelInWaitTime") + rsd("DelOutWaitTime")	'搬出(IN受付→OUT処理完了)
    End If
    rsd.Close
'ADD END HiTS Ver2 By SEIKO N.Ooshige
    conn.Close
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
<!-------------ここからターミナル所要時間画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/terminalt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.07
%>
          </td>
        </tr>
      </table>
      <center>
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
<%
	strNowTime = Year(Now) & "年" & _
		Right("0" & Month(Now), 2) & "月" & _
		Right("0" & Day(Now), 2) & "日" & _
		Right("0" & Hour(Now), 2) & "時" & _
		Right("0" & Minute(Now), 2) & "分現在の情報"

%>
					&nbsp;&nbsp;<%=strNowTime%>
					</font>
				</td>
			</tr>
		</table>

      <table border=0>
        <tr>
          <td align=left colspan="2">
            <br>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                <td nowrap><b>香椎パークポートコンテナターミナル</b></td>
                <td><img src="gif/hr.gif" width="300"></td>
              </tr>
            </table>
          </td></tr>
          <tr><td width="80"></td><td>
<!--
            <center>
			<BR>

			  <table border="0" cellspacing="0" cellpadding="0" width="400">
 				<tr>
				  <td lign=left>
					過去１時間のデータで算出しています。<BR>
					ゲート終了等でデータが得られない場合、値は表示されません。
				  </td>
				</tr>
			  </table>
				<BR>
-->
                <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" width="430">
                  <tr align="center"> 
                    <td nowrap bgcolor="#FFCC33" colspan="3">ターミナル内所要時間</td>
                    <td nowrap bgcolor="#FFCC33" rowspan="2">ゲート前カメラ映像</td>
                  </tr>
                  <tr align="center"> 
                    <td nowrap bgcolor="#FFFFCC"> 搬入のみ </td>
                    <td nowrap align="center"  bgcolor="#FFFFCC"> 搬出のみ </td>
                    <td nowrap align="center"  bgcolor="#FFFFCC"> 搬出入 </td>
                  </tr>
                  <tr align="center"> 
                <td nowrap bgcolor="#FFFFFF" width=90>
<% ' 搬入待ち時間
    If iRecWaitTime>120 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write iRecWaitTime & "分"
    End If
%>
                </td>
                <td nowrap bgcolor="#FFFFFF" width=90>
<% ' 搬出待ち時間
    If iDelWaitTime>120 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write iDelWaitTime & "分"
    End If
%>
                </td>
                <td nowrap bgcolor="#FFFFFF" width=90>
<% ' 搬出入待ち時間
    If iRDWaitTime>120 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write iRDWaitTime & "分"
    End If
%>
                </td>
                <td nowrap align="center"  bgcolor="#FFFFFF">
					<a href="camera.asp"><img src="gif/camera.gif" width="38" height="35" border="0"></a>
				</td>
              </tr>
            </table>
<%'ADD START HiTS Ver.2 By SEIKO N.Ooshige %>
         </td></tr>
         <tr><td align=left colspan="2">
            <br>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                <td nowrap><b>アイランドシティコンテナターミナル</b></td>
                <td><img src="gif/hr.gif" width="300"></td>
              </tr>
            </table>
         </td></tr>
         <tr><td width="80"></td><td>
                  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" width="500">
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFCC33" colspan="2">ターミナル内所要時間</td>
                      <td nowrap bgcolor="#FFCC33" rowspan="2">ゲート前カメラ映像</td>
                    </tr>
                    <tr align="center"> 
                      <td nowrap bgcolor="#FFFFCC"> 搬入(IN受付→作業完了) </td>
                      <td nowrap align="center"  bgcolor="#FFFFCC"> 搬出(IN受付→OUT処理完了) </td>
                    </tr>
                    <tr align="center"> 
                  <td nowrap bgcolor="#FFFFFF">
<% ' IC搬入待ち時間
    If IcInTime<2 or IcInTime>240 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write IcInTime & "分"
    End If
%>
                  </td>
                  <td nowrap bgcolor="#FFFFFF">
<% ' IC搬出待ち時間
    If IcOutTime<2 or IcOutTime>240 Then
        Response.Write "<hr width=80%" & ">"
    Else
        Response.Write IcOutTime & "分"
    End If
%>
                  </td>
                  <td nowrap align="center"  bgcolor="#FFFFFF">
                    <a href="camera.icct.asp"><img src="gif/camera.gif" width="38" height="35" border="0"></a>
                  </td>
                  </tr>
            </table>
         </td></tr>
		 <tr><td>　</td><td></td></tr>
         <tr><td colspan="2">
<%
'			<form>
'			  <table border="0" cellspacing="0" cellpadding="0" width="430">
' 				<tr>
'				  <td lign=left>
'					<input type=button value='表示データの更新' OnClick="JavaScript:location.reload()">
'				  </td>
'				</tr>
'			  </table>
'			</form>
%>
		<form>
			<P>
				過去１時間のデータで算出しています。<BR>
				ゲート終了等でデータが得られない場合、値は表示されません。<BR>
				<input type=button value='表示データの更新' OnClick="JavaScript:location.reload()">
			</P>
		</form>
<%'ADD END HiTS Ver.2 By SEIKO N.Ooshige %>
            <table>
              <tr> 
                <td><img src="gif/botan.gif" width="17" height="17"></td>
                <td nowrap><b>周辺道路状況リンク</b></td>
                <td><img src="gif/hr.gif" width="300"></td>
              </tr>
            </table><br>
            <center>
            <table>
              <tr>
                <td><a href="linklog.asp?link=http://www.fk-tosikou.or.jp" target="_blank">福岡北九州高速道路公社</a></td>
              </tr>
              <tr>
                <td><a href="linklog.asp?link=http://www.jartic.or.jp" target="_blank">（財）日本道路交通情報センター</a></td>
              </tr>
            </table>
            </center>
          </td>
        </tr>
      </table>
      <br>
      </center></td>
  </tr>
  <tr>
    <td valign="bottom">
<%
    DispMenuBar
%>
    </td>
  </tr>
</table>
<!-------------ターミナル所要時間画面終わり--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ターミナル所要時間照会
    WriteLog fs, "8001", "ゲート前映像・混雑状況照会-ゲート内所要時間", "00", ","
%>
