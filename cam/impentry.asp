<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �A���R���e�i�Ɖ�
    WriteLog fs, "2301","�A���R���e�i�Ɖ�","00", ","
%>

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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������Ɖ���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="../gif/impentryt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
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
		<BR>

	  <table border="0" cellpadding="0" cellspacing="0" width="750">
		<tr>
		  <td align="left" nowrap>
			<font size="5" color="#ff6600"><b>�����ӓ�</b></font>
		  </td>
		</tr>
	  </table>

<!-- commented by seiko-denki 2003.07.07
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
		<BR>
<table border=0><tr><td align=left>
      <table>
        <tr>
          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>�L�[���͂̏ꍇ</b></td>
          <td><img src="../gif/hr.gif"></td>
        </tr>
      </table>
<center>
      <table width=500>
        <tr>
          <td colspan="2">�Q�Ƃ������R���e�iNo.�܂��́ABL No.����͂��A�w�A���Ɖ�x�{�^�����N���b�N���ĉ������B�������͂���ꍇ�ɂ�","�ŋ�؂��ē��͂��ĉ������B<br>
          </td>
        </tr>
        <tr>
          <td width="20">&nbsp;</td>
          <td>�����R���e�iNo.���� ��jFYTU2334999,HYKU9882272,DYTU3998821</td>
        </tr>
      </table>
      <form action="impcntnr.asp">
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>�R���e�iNo.</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td>
							<input type=text name=cntnrno size=20 maxlength="100">
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
                <tr> 
                  <td align="center" colspan="2">�܂���
                  </td>
                </tr>
                <tr>
                  <td bgcolor="#000099"><font color="#FFFFFF"><b>BL No.</b></font></td>
                  <td nowrap> 
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td>
							<input type="text" name=blno size=20 maxlength="100">
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
              </table>
              <br>
              <input type=submit value="   �A���Ɖ�   "><br><br>
				�������͂��Ȃ��Łw�A���Ɖ�x�{�^���������ƏƉ�ʂ�
				�T���v����\�����܂��B
      </form>
</center>
		      <table>
		        <tr> 
		          <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
		          <td nowrap><b>CSV�t�@�C�����͂̏ꍇ</b></td>
		          <td><img src="../gif/hr.gif"></td>
		        </tr>
		      </table>
<center>
		      <table border="0" cellspacing="1" cellpadding="2">
		        <tr> 
		          <td> 
		            <p>�����������t�@�C���]������ꍇ�͂������N���b�N</p>
		          </td>
		          <td>�c</td>
		          <td><a href="impcsv.asp">CSV�t�@�C���]��</a></td>
		        </tr>
		        <tr> 
		          <td>CSV�t�@�C���]���ɂ��Ă̐����͂������N���b�N</td>
		          <td>�c</td>
		          <td><a href="help02.asp">�w���v</a></td>
		        </tr>
		      </table>

			  <br>
<!--
			  <form>
				<input type="button" value="����CT�EICCT" style="width: 150px" onclick="javascript:location.href='../impentry.asp'">
			  </form>
-->
</center>

		  </td>
		</tr>
	  </table>

      <br>
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
<!-------------�Ɖ��ʏI���--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>
