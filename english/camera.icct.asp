<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' ゲート周辺簡易地図
    WriteLog fs, "8006", "ゲート前映像・混雑状況照会-ICCTゲート周辺簡易地図", "00", ","
%>

<html>
<head>
<title>ゲート周辺簡易地図</title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
  <td valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan="2"><img src="gif/terminalt.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
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
        <table width="95%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td align="right">
              <font color="#333333" size="-1"><%=strRoute%></font>
            </td>
 </tr>
</table>
End of comment by seiko-denki 2003.07.18 -->


<%
    Dim fs, f, strUpdateTime, strFileName, strPath
    Set fs = CreateObject("Scripting.FileSystemObject")
    strFileName="./camera/sam-gate.icct.jpg"
	strPath = Server.MapPath(strFileName)
    Set f = fs.GetFile(strPath)
	dateTimeTmp = f.DateLastModified
	strUpdateTime = Year(dateTimeTmp) & "年" & _
		Right("0" & Month(dateTimeTmp), 2) & "月" & _
		Right("0" & Day(dateTimeTmp), 2) & "日" & _
		Right("0" & Hour(dateTimeTmp), 2) & "時" & _
		Right("0" & Minute(dateTimeTmp), 2) & "分現在の情報"
%>

		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
					&nbsp;&nbsp;<%=strUpdateTime%>
					</font>
				</td>
			</tr>
		</table>


<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17"></td>
            <td nowrap><b>ゲート周辺簡易地図</b></td>
            <td><img src="gif/hr.gif" width="400" height="3"></td>
          </tr>
        </table>
<table border="0"><tr>
            <td align="left"><br>
              <br>
<center>
                <img src="gif/terminalmap.icct.gif" width="440" height="252" usemap="#TarminalMap" border="0"> 
                <br>
                　<br>

<table border=0 cellpadding=0 cellspacing=0>
 <tr>
  <td align="center" valign=middle>

	<table border=5 cellpadding=0 cellspacing=0 bgcolor="#6666ff" bordercolor="#ffffff">
	 <tr>
	  <td align="center" valign=middle height=20>
		<font color="#ffffff"><b>ゲート前映像</b></font>
	  </td>
	 </tr>
	 <tr>
	  <td align=center valign=middle><a href="photogate.icct.asp"><img src="camera/sam-gate.icct.jpg" border="0" width="120" height="82"></a></td>
	 </tr>
	</table>

  </td>
 </tr>
 <tr>
   <td align=center valign=bottom>
	<b>写真をクリックすると拡大できます。</b>
   </td>
 </tr>
</table>
<P>
<table border="0">
 <tr>
   <td align="left">
<form>
				キャッシュ機能により古い画像が表示されることがありますので、<br>
				最新の表示にするには下のボタンをクリックしてください。<br>
				<input type=button value="表示データの更新" onClick="JavaScript:window.location.reload()">
			</form>
   </td>
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
<!-------------登録画面終わり--------------------------->
<%
    DispMenuBarBack "terminal.asp"
%>
<map name="TarminalMap"> 
  <area shape="rect" coords="70,183,162,209" href="photogate.icct.asp">
</map>
</body>
</html>