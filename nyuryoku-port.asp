<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-in1.asp"

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �����w��̂Ȃ��Ƃ�
        strFileName="test.csv"
    End If
    strFileName="./temp/" & strFileName

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �{�D���Ê�{���̎擾
    If Not ti.AtEndOfStream Then
        anyTmp=Split(ti.ReadLine,",")
    End If

    ' �ڍו\���s�̃f�[�^���̎擾
    If Not ti.AtEndOfStream Then
        iCount=CInt(ti.ReadLine)
    End If
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<meta http-equiv="Pragma" content="no-cache">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>

<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!-------------��������ꗗ���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/nyuryoku-s.gif" width="506" height="73"></td>
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
<!-- commented by seiko-denki 2003.07.18--->
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
<!-- End of Addition by seiko-denki 2003.07.18--->
		<BR>
		<BR>
		<BR>
<table border=0><tr><td align=left>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>�{�D���Èꗗ</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <table>
        <tr>
          <td align=left>
            <table border=1 cellpadding="3" cellspacing="1">
              <tr> 
                <td bgcolor="#000099" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�D��</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' �D�Ж��̕\��
    Response.Write anyTmp(1)
%>
                </td>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�D��</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' �D���̕\��
    Response.Write anyTmp(3)
%>
                </td>


                <td bgcolor="#000099" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>Voyage No.</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' ���q�̕\��
    If anyTmp(5)=anyTmp(6) Then
        Response.Write anyTmp(5)
    Else
        Response.Write anyTmp(5) & "/" & anyTmp(6)
    End If
%>
                </td>
                <td bgcolor="#003399" background="gif/tableback.gif" nowrap><font color="#FFFFFF"><b>�R�[���T�C��</b></font></td>
                <td bgcolor="#FFFFFF" nowrap>
<%
    ' �R�[���T�C���̕\��
    Response.Write anyTmp(2)
%>
                </td>
              </tr>
            </table>
			<BR>

			<table border=0 cellpadding=1><tr><td width=30></td>
			<td nowrap>
			�f�[�^���X�V����ꍇ�͑ΏۂƂȂ�`����I�����ĉ������B<BR>
			�V�K�|�[�g��ǉ�����ꍇ�͐V�K�|�[�g��I�����ĉ������B
			</td></tr></table>

            <table>
              <tr>
                <td>
                  <table border="1" cellspacing="1" cellpadding="3">
                    <tr bgcolor="#FFCC33">
                      <td nowrap align=center valign=middle><br></td>
                      <td nowrap align=center valign=middle>�`��</td>
                      <td nowrap align=center valign=middle>���ݗ\�莞��</font></td>
                      <td nowrap align=center valign=middle>���݊�������</font></td>
                      <td nowrap align=center valign=middle>���݊�������</font></td>
                      <td nowrap align=center valign=middle>���� Long Schedule</font></td>
                      <td nowrap align=center valign=middle>���� Long Schedule</font></td>
                    </tr>
<!-- ��������f�[�^�J��Ԃ� -->
<%
    LineNo=1
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        Response.Write "<tr bgcolor='#FFFFFF'>"
        Response.Write "<td align=center>" & LineNo & "</td>"
        Response.Write "<td nowrap align=center valign=middle>"
        Response.Write "<a href='nyuryoku-sch.asp?line=" & LineNo & "'>"
        Response.Write anyTmp(1) & "</a></td>"
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(2)<>"" Then
            Response.Write anyTmp(2) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(3)<>"" Then
            Response.Write anyTmp(3) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(5)<>"" Then
            Response.Write anyTmp(5) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(6)<>"" Then
            Response.Write Left(anyTmp(6),10) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "<td nowrap align=center valign=middle>"
        If anyTmp(7)<>"" Then
            Response.Write Left(anyTmp(7),10) & "</td>"
        Else
            Response.Write "<hr width=80%" & "></td>"
        End If
        Response.Write "</tr>"
        LineNo=LineNo+1
    Loop
    ti.Close
%>
<!-- �����܂� -->
                    <tr bgcolor="#FFFFFF"> 
                      <td><br></td>
                      <td nowrap align=center valign=middle>
                        <a href="nyuryoku-new.asp">�V�K�|�[�g</a>
                      </td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                      <td nowrap align=center valign=middle><hr width=80%></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <br>
            <center>
            <form action="nyuryoku-vsl.asp"><input type="submit" name="submit" value="  ��  �M  ">
            </form>
            </center>
          </td>
        </tr>
      </table>
</td></tr></table>

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
<!-------------�ꗗ��ʏI���--------------------------->
<%
    DispMenuBarBack "nyuryoku-in1.asp"
%>
</body>
</html>

<%
    ' �{�D���Ó��͈ꗗ
    WriteLog fs, "3003","�D�Ё^�^�[�~�i������-�{�D���Èꗗ","00", ","
%>
