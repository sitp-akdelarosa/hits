<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'	�������o�V�X�e���y�C�ݗp�z	�ύX,�폜�p���
%>

<%
	' �Z�b�V�����̃`�F�b�N
	CheckLogin "sokuji.asp"

	' �C�݃R�[�h
	sForwarder = Trim(Session.Contents("userid"))

	' �V�K�ǉ���(2) or �V�K(1) or �X�V(0)
	Dim bKind
	bKind = Request.QueryString("kind")

	If bKind=0 Then
		Session.Contents("kind") = 0
	ElseIf bKind=1 Then
		Session.Contents("kind") = 1
	ElseIf bKind=2 Then
		Session.Contents("kind") = 2
	End If

	If bKind = 0 Then
		Dim sShipper,sShipLine,sVslCode,sBL,sCont,sReject,sRecschTime,iLineNo
		sShipper 	= Request.form("shipper")
		sShipLine 	= Request.form("shipline")
		sVslCode 	= Request.form("vslcode")
		sOpe 		= Request.form("ope")
		sOpeTel		= Request.form("opetel")
		sBL 		= Request.form("bl")
		sCont 		= Request.form("cont")
		sReject 	= Request.form("reject")
		sRecschTime = Request.form("recschtime")
		iLineNo		= Request.form("lineno")
	End If

%>

<html>
<head>
<title>�������o�\���݁i�C�݁j</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
	function checkFormValue(){
		if(!checkBlank(getFormValue(0))){ return showAlert("�׎�R�[�h",true); }
		if(!checkBlank(getFormValue(1))){ return showAlert("�D�ЃR�[�h",true); }
		if(!checkBlank(getFormValue(2))){ return showAlert("�D���R�[�h",true); }
		if(!checkBlank(getFormValue(3)) && !checkBlank(getFormValue(4))){ return showAlert("BL No.�܂��̓R���e�iNo.",true); }
		if(checkBlank(getFormValue(3)) && checkBlank(getFormValue(4))){ return showAlert("BL No.�ƃR���e�iNo.",false); }
		return true;
	}
	function getFormValue(iNum){
		formvalue = window.document.input.elements[iNum].value;
		return formvalue;
	}

	function checkBlank(formvalue){
		if(formvalue == ""){ return false; }
		return true;
	}
	function showAlert(strAlert,bKind){
		if(bKind){
			window.alert(strAlert + "�������͂ł��B");
		} else {
			window.alert(strAlert + "�́A�ǂ��炩�������͂��ĉ������B");
		}
		return false;
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
          <td rowspan=2><img src="gif/sokuji1t.gif" width="506" height="73"></td>
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
          <td> 

	        <table>
	          <tr>
	            <td><img src="gif/botan.gif" width="17" height="17"></td>
	            <td nowrap><b>�i�C�ݗp�j�������o�\����</b></td>
	            <td><img src="gif/hr.gif"></td>
	          </tr>
	        </table>

              <center>
            <br>
			�������o�Ώۉݕ��ɂ��āA���̊e���ڂ���͂��ĉ������B

            <form method=post name="input" action="sokuji-kaika-exec.asp">
			<table border=0 cellpadding=0 cellspacing=0>
			  <tr>
				<td nowrap align=left>


              <table border="1" cellspacing="2" cellpadding="2" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" width=120 align=center valign=middle>
                    <font color="#FFFFFF"><b>�׎�R�[�h</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180 height=25 valign=middle>
<% If bKind=1 Then %>
							<input type=text name=shipper value="<%=sShipper%>" size=7 maxlength=5>
<% Else %>
							<font>&nbsp;<%=sShipper%></font>
							<input type=hidden name="shipper" value="<%=sShipper%>">
<% End If %>
						</td>
						<td align=left valign=middle nowrap>
<% If bKind=1 Then %>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
<% End If %>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�D�ЃR�[�h</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180 height=25 valign=middle>
<% If bKind=1 Then %>
							<input type=text name=shipline value="<%=sShipLine%>" size=7 maxlength=5>
<% Else %>
							<font>&nbsp;<%=sShipLine%></font>
							<input type=hidden name="shipline" value="<%=sShipLine%>">
<% End If %>
						</td>
						<td align=left valign=middle nowrap>
<% If bKind=1 Then %>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
<% End If %>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�D���R�[�h</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180 height=25 valign=middle>
<% If bKind=1 Then %>
							<input type=text name=vslcode value="<%=sVslCode%>" size=9 maxlength=7>
<% Else %>
							<font>&nbsp;<%=sVslCode%></font>
							<input type=hidden name="vslcode" value="<%=sVslCode%>">
<% End If %>
						</td>
						<td align=left valign=middle nowrap>
<% If bKind=1 Then %>
							<font size=1 color="#ee2200">[ �K�{���� ]</font><BR>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
<% End If %>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

			  </table>

			  <BR><BR>

				<center>BL No.�܂��́A�R���e�iNo.�̂ǂ��炩����͂��đ��M�{�^���������ĉ������B</center>
				<BR>

              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">

                <tr> 
                  <td bgcolor="#000099" width=120 align=center valign=middle>
                    <font color="#FFFFFF"><b>BL No.�̏ꍇ</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180>
							<input type=text name=bl value="<%=sBL%>" size=22 maxlength=20>
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
                    <font color="#FFFFFF"><b>�R���e�iNo.�̏ꍇ</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=180>
							<input type=text name=cont value="<%=sCont%>" size=14 maxlength=12>
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#2288ff">[ ���p�p�� ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>

              </table>
              <br>
				<center>

<% If bKind=0 Then %>
				<input type=hidden name="blold" value="<%=sBL%>">
				<input type=hidden name="contold" value="<%=sCont%>">
<% End If %>

				<input type=hidden name="ope" value="<%=sOpe%>">
			  <input type=hidden name="opetel" value="<%=sOpeTel%>">
			  <input type=hidden name="reject" value="<%=sReject%>">
			  <input type=hidden name="recschTime" value="<%=sRecschTime%>">
			  <input type=hidden name="lineno" value="<%=iLineNo%>">
              <input type=submit name="send" value=" ��  �M " onClick="return checkFormValue()">
              <input type=submit name="stop" value=" �I  �� ">

<% If bKind<>1 Then %>

              <input type=submit name="del" value=" ��  �� ">

<% End If %>

				</center>
				</td>
			  </tr>
			</table>
            </form>
              </center>
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
    DispMenuBarBack "sokuji-kaika-list.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log�쐬
    If bKind=0 Then
    	WriteLog fs,"7002", "�������o�V�X�e��-�C�ݗp�\����", "02", ","
	Else
	    WriteLog fs,"7002", "�������o�V�X�e��-�C�ݗp�\����", "01", ","
	End If
%>
