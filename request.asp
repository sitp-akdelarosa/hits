<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<META name="GENERATOR" content="IBM HomePage Builder 2001 V5.0.0 for Windows">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function ClickSend() {

	if (document.con.name1.value == "") {	
		window.alert("��Ж��������͂ł��B");
		return false;
	}

	if (document.con.address1.value == "") {	
		window.alert("�Z���������͂ł��B");
		return false;
	}

}

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->

<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/requestt.gif" width="506" height="73"></td>
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
		<BR>
		<BR>
		<BR>
   					<form NAME="con" action="request-syori.asp" method=post onSubmit="return ClickSend()">
      <table width="500" cellpadding="0">
					<tr>
						<td bordercolor="#FFFFFF">HiTS V3�������p�������肪�Ƃ��������܂��B
            <BR>
<BR>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� ���ⓙ���������܂�����<a href="mailto:mrhits@hits-h.com">E-mail</a>�ł��⍇���������B <br>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �p���`�W�̃y�[�W��<a href="qa/index.html">������</a>���炲�Q�ƂɂȂ�܂��B<BR>
<BR>
							�Ȃ��A���}���̏ꍇ�͉��L�ɂ��₢���킹�������B<BR><BR>

							<table border=0 cellpadding=1 cellspacing=1>
							  <tr>
								<td colspan=5 align=left>
								&nbsp;�y �d�b�ł̂��₢���킹�� �z
								</td>
							  </tr>

							  <tr><td colspan=5 height=2></td></tr>

							  <tr>
								<td width=20 rowspan=3><BR></td>
								<td nowrap align=left valign=top colspan=4>�E�����`����IT�V�X�e���̉^�p�Ɋւ��鎖<BR>
								</td>
							  </tr>
							  <tr>
								<td width=15 rowspan=2><BR></td>
								<td nowrap align=left valign=top colspan=2>�����`�ӓ��������</td>
								<td align=left nowrap>�S���F�ؖ{</td>
							  </tr>
							  <tr>
								<td width=15><BR></td>
								<td nowrap align=left valign=top>TEL 092-663-3021<BR>
								</td>
								<td><BR></td>
							  </tr>

							  <tr><td colspan=5 height=5></td></tr>

							  <tr>
								<td width=20 rowspan=3><BR></td>
								<td nowrap align=left valign=top colspan=4>�E�����`����IT�V�X�e���̊J���Ɋւ��鎖<BR>
								</td>
							  </tr>
							  <tr>
								<td width=15 rowspan=2><BR></td>
								<td nowrap align=left valign=top colspan=2>�����s�`�p�ǁ@�`�p�U�����@��������</td>
								<!-- <td align=left nowrap>�S���F�J��</td> -->
							  </tr>
							  <tr>
								<td width=15><BR></td>
								<td nowrap align=left valign=top>TEL 092-282-7108<BR>
                  </td>
								<td><BR></td>
							  </tr>

							  <tr><td colspan=5 height=5></td></tr>

							  <tr>
								<td width=20 rowspan=3><BR></td>
								<td nowrap align=left valign=top colspan=4><BR>
								</td>
							  </tr>
							  <tr>
								<td width=15 rowspan=2><BR></td>
								<td nowrap align=left valign=top colspan=2></td>
								<td align=left nowrap></td>
							  </tr>
							  <tr>
								<td width=15></td>
								<td nowrap align=left valign=top></td>
								<td></td>
							  </tr>

							</table>
						</td>
					</tr>
				</table> 
<BR>
					
<%
    DispMenuBar
%>
		</FORM></CENTER></td>
	</tr>
</table>
<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>
<%

    WriteLog fs, "9001", "���p�҃A���P�[�g�EQ&A", "00", ","
%>
