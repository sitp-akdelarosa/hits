<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	'	�y�C�ݓ��́z	�V�K���́A�b�r�u�]���A�X�V�I�����
%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "pickselect.asp"

    strUserKind=Session.Contents("userkind")
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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------�������烍�O�C�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=94%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
<%
	If strUserKind="�C��" Then
		titlegif = "pickkat"
	Else
		titlegif = "picknit"
	End If
%>
          <td rowspan=2><img src="gif/<%=titlegif%>.gif" width="506" height="73"></td>
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

<table border=0><tr><td align=left>

      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>�V�K����</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>

	<center>
      <br>
      <table>
        <tr>
          <td nowrap align=center>
			��R���s�b�N�A�b�v�˗���V�K�œ��͂���ꍇ�͂������N���b�N ��� <a href="pickexp-input.asp?kind=1">�V�K����</a>
		  </td>
		</tr>
	  </table>
	</center>

      <br><BR>
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
            <p>��R���s�b�N�A�b�v�˗����t�@�C���]������ꍇ�͂������N���b�N</p>
          </td>
          <td>�c</td>
          <td><a href="pickexp-csv.asp">CSV�t�@�C���]��</a></td>
        </tr>
        <tr> 
          <td>CSV�t�@�C���]���ɂ��Ă̐����͂������N���b�N</td>
          <td>�c</td>
<%
	If strUserKind="�C��" Then
		ihelpnum = 24
	Else
		ihelpnum = 25
	End If
%>
          <td><a href="help<%=ihelpnum%>.asp">�w���v</a></td>
        </tr>
      </table>
	</center>

	  <BR><BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17"></td>
          <td nowrap><b>���͍σf�[�^�̍X�V</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>

	<center>
      <table>
        <tr>
          <td nowrap>
            <form method=post action="pickexp-dblist.asp">
              <center>
				<table border=0 cellpadding=0>
				  <tr>
					<td align=left>
				�i�荞�݂��s���ꍇ�́A���L�t�H�[���ɓK���Ȓl����͂��Ă���<BR>�X�V�Ώۈꗗ�{�^���������ĉ������B
				<BR><BR>��������͂��Ȃ��ŏƉ�����s����ƁA�S���\������܂��B
					</td>
				  </tr>
				</table>
				<BR>
              <table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
                <tr> 
<% 
	If strUserKind="�C��" Then
		UserSerchCodeName = "�׎�"
	Else
		UserSerchCodeName = "�C��"
	End If
%>
                 <td bgcolor="#000099" nowrap align=center valign=middle> <font color="#FFFFFF"><b><%=UserSerchCodeName%>�R�[�h</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=100>
							<input type=text name=suser size=7 maxlength=5>
						</td>
						<td align=right valign=middle nowrap>
							<font size=1 color="#2288ff">[���p�p��]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�׎�Ǘ��ԍ�</b></font>
                  </td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td width=100>
							<input type=text name=suserno size=12 maxlength=10>
						</td>
						<td align=right valign=middle nowrap>
							<font size=1 color="#2288ff">[���p�p��]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
              </table>
              <br>
                <input type=submit value="�X�V�Ώۈꗗ"> 
              </center>
            </form>
		  </td>
		</tr>
	  </table>
	</center>
          </td>
        </tr>
      </table>

</td></tr></table>
    </td>
  </tr>
  <tr>
    <td valign="bottom"> 
<%
    DispMenuBar
%>
<!-------------���O�C����ʏI���--------------------------->
    </td>
  </tr>
</table>
<%
    DispMenuBarBack "pickselect.asp"
%>
</body>
</html>

<%
    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
	' Log�쐬
	If strUserKind="�C��" Then
		iWrkNum = "01"
	Else
		iWrkNum = "02"
	End If
    WriteLog fs, "a110","��R���s�b�N�A�b�v�V�X�e��-��R���s�b�N�A�b�v�˗�",iWrkNum, ","
%>
