<%@Language="VBScript" %>

<!--#include file="Common.inc"-->


<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/impentryt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.07
'	DisplayCodeListButton
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
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>CSV file transfer</b></td>
          <td><img src="gif/hr.gif" hspace="3"></td>
        </tr>
      </table>
      <table width="450">
        <tr> 
          <td>
            <dl> 
              <dt><font color="#000066" size="+1">[for import container]</font><br>
              <dd nowrap>Please select the CSV file that inputs the search condition,<BR>
              and click 'Transmission' button.(File�@format= One a line)</dl></td></tr>
      </table>
    <form action="impcsvin.asp?kind=cntnr" method="POST" enctype="multipart/form-data">
      <table border="1" cellspacing="2" cellpadding="3" bgcolor="#FFFFFF">
        <tr> 
          <td bgcolor="#000099" nowrap align="center" background="gif/tableback.gif" colspan="2"><font color="#FFFFFF"><b>At the CSV file that enumerates container No.</b></font>
          </td>
        </tr>
        <tr>
          <td bgcolor="#000099" nowrap>
            <font color="#FFFFFF"><b>CSV file</b></font>
          </td>
          <td nowrap> 
            <input type=file name=csvfile size=50 accept="text/css">
          </td>
        </tr>
      </table>
		<BR>
      <input type=submit value="transmit">
      <br><br>
    </form>
    <br><br>
    <form action="impcsvin.asp?kind=blno" method="POST" enctype="multipart/form-data">
      <table border="1" cellspacing="2" cellpadding="3" bgcolor="#FFFFFF">
        <tr> 
          <td bgcolor="#000099" nowrap align="center" background="gif/tableback.gif" colspan="2"><font color="#FFFFFF"><b>At the CSV file that enumerates BL No.</b></font>
          </td>
        </tr>
        <tr>
          <td bgcolor="#000099" nowrap>
            <font color="#FFFFFF"><b>CSV file</b></font>
          </td>
          <td nowrap> 
            <input type=file name=csvfile size=50 accept="text/css">
          </td>
        </tr>
      </table>
		<BR>
      <input type=submit value="transmit">
    </form>
    <br>
    <br>
      </center>
  </td></tr> 
  <tr>
    <td valign="bottom">
<%
    DispMenuBar
%>
    </td>
  </tr>
</table> 
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "impentry.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    WriteLog fs, "2003","�A���R���e�i�Ɖ�-CSV�t�@�C���]��","00", ","
%>
