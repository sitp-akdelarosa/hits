<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	'
	'	�y�R���e�i�����́z	���͉��
	'
%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"
	' �C�݃R�[�h
	sSosin = Trim(Session.Contents("userid"))

	' �V�K(1) or �X�V(0)
    bKind = Request.QueryString("kind")
	Dim sUser,sUserNo,sVslCode,sVoyCtrl,sBooking,sCont,sSeal,sCargoWeight,sContWeight,sRifer,sDanger
	sUser 		= Request.form("user")
	sUserNo 	= Request.form("userno")
	sVslCode 	= Request.form("vslcode")
	sVoyCtrl 	= Request.form("voyctrl")
	sBooking 	= Request.form("booking")
	If bKind=0 Then
		sCont 		= Request.form("cont")
		sSeal 		= Request.form("seal")
		sCargoWeight= Request.form("cargow")
		sContWeight	= Request.form("contw")
		sRifer 		= Request.form("rifer")
		sDanger 	= Request.form("danger")
	End If
	iLineNo		= Request.form("lineno")

%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

	function checkFormValue(){
		contvalue = window.document.input.cont.value;
		if(contvalue == ""){
			window.alert("�R���e�iNo.�������͂ł��B");
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

<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------�������烍�O�C�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika5t.gif" width="506" height="73"></td>
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
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>�X�V������</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>
      <table>
        <tr>
          <td nowrap align=center>
				�A�o�R���e�i�ɂ��āA�ȉ��̍��ڂ���͂��đ��M�������ĉ������B
            <form method=post name="input" action="ms-kaika-expcontinfo-exec.asp">
				<input type=hidden name="kind" value="<%=bKind%>">
				<input type=hidden name="lineno" value="<%=iLineNo%>">
              <center>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�D��</b></font>
                  </td>
                  <td nowrap>
                    <%=sVslCode%>
					<input type=hidden name="vslcode" value="<%=sVslCode%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Voyage No.</b></font>
                  </td>
                  <td nowrap>
                    <%=sVoyCtrl%>
					<input type=hidden name="voyctrl" value="<%=sVoyCtrl%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
					 <font color="#FFFFFF"><b>�׎�R�[�h</b></font>
				  </td>
                  <td nowrap>
                    <%=sUser%>
					<input type=hidden name="user" value="<%=sUser%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�׎�Ǘ��ԍ�</b></font>
                  </td>
                  <td nowrap>
                    <%=sUserNo%>
					<input type=hidden name="userno" value="<%=sUserNo%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Booking No.</b></font>
                  </td>
                  <td nowrap>
                    <%=sBooking%>
					<input type=hidden name="booking" value="<%=sBooking%>">
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�R���e�iNo.</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=cont value="<%=sCont%>" size=14 maxlength=12>
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
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�V�[��No.</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=seal value="<%=sSeal%>" size=17 maxlength=15>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                    
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�ݕ��d��</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=cargow value="<%=sCargoWeight%>" size=5 maxlength=4 onblur="checknum(document.input.cargow)">�it�j
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
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>���d��</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=150>
							<input type=text name=contw value="<%=sContWeight%>" size=5 maxlength=4 onblur="checknum(document.input.contw)">�it�j
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
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>���[�t�@�[</b></font>
                  </td>
                  <td nowrap>
<%	
	Dim strRifKind
	If bKind=0 And sRifer="R" Then
		strRifKind = "checked"
	End If
%>
					<input type=checkbox name=rifer <%=strRifKind%>>
					<font size=-1>���[�t�@�[�̏ꍇ�`�F�b�N���ĉ������B</font>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�댯��</b></font>
                  </td>
                  <td nowrap>
<%	
	Dim strDngKind
	If bKind=0 And sDanger="H" Then
		strDngKind = "checked"
	End If
%>
					<input type=checkbox name=danger <%=strDngKind%>>
					<font size=-1>�댯���̏ꍇ�`�F�b�N���ĉ������B<sup>�i���j</sup></font>
                  </td>
                </tr>

              </table>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font size=-1>�i���j ���h�@�Ɋւ��댯���̏ꍇ�̂݃`�F�b�N���ĉ������B</font>
              <br><BR>
                <input type=submit name="send" value=" ��  �M " onClick="return checkFormValue()">
                <input type=button value=" ��  �~ " onClick="JavaScript:window.history.back()">

				</center>
              </center>
            </form>
<%
            If bError Then
                ' �G���[���b�Z�[�W�̕\��
                DispErrorMessage strError
            End If
%>
          </td>
        </tr>
      </table>
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
<!-------------���O�C����ʏI���--------------------------->
<%
    DispMenuBarBack "ms-kaika-expcontinfo.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log�쐬
    WriteLog fs, "4106","�C�ݓ��͗A�o�R���e�i���-������","00", ","
%>
