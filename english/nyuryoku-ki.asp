<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
     '�Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-ki.asp"

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function ClickSend() {
	if (ChkSend(document.con.ContNo.value, 
				document.con.SealNo.value, 
				document.con.Jyuryo.value,
				document.con.SoJyuryo.value)) { 
		return true;
	}
	return false;
}
// ���̓`�F�b�N
function ChkSend(sContNo, sSealNo, sJyuryo, sSoJyuryo) {

	if (sContNo == "") {	/* �R���e�iNo.�����̓`�F�b�N */
			window.alert("�R���e�iNo.�������͂ł��B");
			return false;
	}

	if (sSealNo == "" && sJyuryo == "" && sSoJyuryo == "") {	/* �V�[��No.�E�d�ʖ����̓`�F�b�N */
			window.alert("�ڍ׏�����͂��ĉ������B");
			return false;
	}
	return true;
}

// ���l�`�F�b�N
function checknum(etext)
{
	if (etext.value == "")
		return false;

	if (isNaN(etext.value)) {
		alert("���l����͂��ĉ������B");
		etext.focus();
		etext.select();
		return false;
	}

	fTemp=parseFloat(etext.value)
    if (fTemp>99.9) {
		alert("99.9Ton�ȉ��̐��l����͂��ĉ������B");
		etext.focus();
		etext.select();
		return false;
	}

	return true;
}

<!--
function gotoURL(){
    var gotoUrl=document.con.select.options[document.con.select.selectedIndex].value
    document.location.href=gotoUrl 
}
//-->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
	<td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika1-2t.gif" width="506" height="73"></td>
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
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
<table border=0 cellpadding=0 cellspacing=0><tr><td align=left>
		<table>
          <tr> 
            <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
            <td nowrap><b>�R���e�i������</b></td>
            <td><img src="gif/hr.gif"></td>
          </tr>
		</table>
<center>
	    <table>
	      <tr>
	        <td>���L�̍��ڂ���͂̏�A���M�{�^�����N���b�N���ĉ������B</td>
          </tr>
		</table>
		  <FORM NAME="con" METHOD="post" action="nyuryoku-ki-syori.asp" onSubmit="return ClickSend()">
                <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                      �R���e�iNo.</font></b></td>
                    <td> 
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="ContNo" size="20" maxlength="12">
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
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">�V�[��No.</font></b></td>
                    <td> 
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="SealNo" size="20" maxlength="15">
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#2288ff">[ ���p�p�� ]</font>
							</td>
						  </tr>
						</table>
                      
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">�ݕ��d��</font></b></td>
                    <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="Jyuryo" size="6"  maxlength="4" onblur="checknum(document.con.Jyuryo)">�it�j
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#2288ff">[ ���p���l ]</font>
							</td>
						  </tr>
						</table>
                      
						&nbsp;&nbsp;&nbsp;<font size="-1">�����_�ȉ�1���܂ŗL��&nbsp;&nbsp;�i��j10.2</font>
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">���d��</font></b></td>
                    <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="SoJyuryo" size="6"  maxlength="4" onblur="checknum(document.con.SoJyuryo)">�it�j
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#2288ff">[ ���p���l ]</font>
							</td>
						  </tr>
						</table>
                      
						&nbsp;&nbsp;&nbsp;<font size="-1">�����_�ȉ�1���܂ŗL��&nbsp;&nbsp;�i��j10.2</font>
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">���[�t�@�[</font></b></td>
                    <td>
                      <input type=checkbox name="rf"><font size=-1>���[�t�@�[�̏ꍇ�̓`�F�b�N���ĉ������B</font>
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">�댯��</font></b></td>
                    <td>
                      <input type=checkbox name="dg"><font size=-1>�댯���̏ꍇ�̓`�F�b�N���ĉ������B<sup>�i���j</sup></font>
                    </td>
                  </tr>
                </table>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font size=-1>�i���j ���h�@�Ɋւ��댯���̏ꍇ�̂݃`�F�b�N���ĉ������B</font>
                <br><BR>
                <input type=submit value=" ��  �M " name="���Z�b�g">
        </form>
</center>

          <br>
          <table>
            <tr> 
              <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
              <td nowrap><b>CSV�t�@�C���]��</b></td>
              <td><img src="gif/hr.gif"></td>
            </tr>
          </table>

<center>
		  <table border="0" cellspacing="1" cellpadding="2">
          <tr> 
              <td> 
                <p>�����t�@�C���]������ꍇ�͂������N���b�N</p>
              </td>
              <td>�c</td>
              <td><a href="nyuryoku-kcsv.asp">CSV�t�@�C���]��</a></td>
            </tr>
            <tr> 
              <td>CSV�t�@�C���]���ɂ��Ă̐����͂������N���b�N</td>
              <td>�c</td>
              <td><a href="help08.asp">�w���v</a></td>
            </tr>
          </table>
</center>
          <br>
          �@<br>
</td></tr></table>
				</center>
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
 </td>
 </tr>
 </table>

<!-------------�o�^��ʏI���--------------------------->
<%
    DispMenuBarBack "nyuryoku-kaika.asp"
%>
</body>
</html>
<%
    ' �C�ݓ��̓V�[��No.�A�d�ʓ���
    WriteLog fs, "4002","�C�ݓ��̓V�[��No.�E�d�ʓ���", "00",","
%>
