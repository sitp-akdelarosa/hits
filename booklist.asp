<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "EXPORT", "expentry.asp"

	Dim strBookingNo
	strBookingNo = ""
'2006/03/06 add-s h.matsuda
  dim ShipLine,ShoriMode
  ShoriMode = ""
  ShipLine = ""
'2006/03/06 add-e h.matsuda

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "expentry.asp"             '�A�o�R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="./temp/" & strFileName

    ' �A�o�R���e�i�Ɖ�X�g�\��
    WriteLog fs, "1010","�u�b�L���O���Ɖ�-�u�b�L���O���ꗗ","00", ","

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)
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
<!-------------��������ꗗ���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/bookingt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.17
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.17
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.17 -->
		<table width=95% cellpadding=3>
			<tr>
				<td align=right>
					<font color="#224599">
					&nbsp;&nbsp;<%=GetUpdateTime(fs)%>
					</font>
				</td>
			</tr>
		</table>
      <table>
        <tr>
          <td> 
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�u�b�L���O���ꗗ</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>

            <table border="0" cellspacing="2" cellpadding="1">
              <tr> 
                <td width="15"><BR></td>
                <td><font color="#000000" size="-1">�i��1�j96=HC</font></td>
                <td width="15"><BR></td>
                <td><font color="#000000" size="-1">�i��2) �N���b�N�Ńs�b�N�A�b�v�σR���e�iNo.��\��</font></td>
              </tr>
            </table>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 

                <td nowrap>Booking No.</td>
                <td nowrap>�D��</td>
                <td nowrap>�D��</td>
				<td nowrap>Voyage No.</td>
                <td nowrap>�d���`</td>
				<td nowrap>��R�����o�ꏊ</td>
<!-- 2008.01.12 NICS START -->
				<td nowrap>CY�J�b�g</td>
<!-- 2008.01.12 NICS END -->
				<td nowrap>�T�C�Y</td>
				<td nowrap>�^�C�v</td>
				<td nowrap>����<font size="-1"><sup>(��1)</sup></font></td>
<!-- I20040223 S -->
				<td nowrap>�ގ�</td>
<!-- I20040223 E -->
				<td nowrap>�\��<BR>�{��</td>
				<td nowrap>�s�b�N�A�b�v��<BR>�{��<font size="-1"><sup>(��2)</sup></font></td>
              </tr>
<!-- ��������f�[�^�J��Ԃ� -->
<% ' �\���t�@�C���̃��R�[�h������ԌJ��Ԃ�
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
%>
              <tr bgcolor="#FFFFFF"> 
				<td nowrap align=center valign=middle>
<% ' Booking No
        If strBooking<>anyTmp(1) Then
            Response.Write anyTmp(1)
            strBooking=anyTmp(1)

			'Reload�p
			If strBookingNo="" Then
				strBookingNo = anyTmp(1)
				'2006/03/06 add-s h.matsuda �u�b�L���O�d�������ɑΉ�
				  if ubound(anyTmp)>14 then
					if trim(anyTmp(15))="ShoriMode=EMoutInf" then
						ShipLine = trim(anyTmp(14))
						ShoriMode = trim(mid(anyTmp(15),11))
					end if
				  end if
				'2006/03/06 add-e h.matsuda
			Else
				strBookingNo = strBookingNo & "," & anyTmp(1)
			End If
        Else
            Response.Write "<br>"
        End If
%>
				</td>
                <td nowrap align=center valign=middle>
<% ' �D��
        If anyTmp(2)<>"" Then
		    Response.Write anyTmp(2)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �D��
        If anyTmp(3)<>"" Then
		    Response.Write anyTmp(3)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' Voyage No.
        If anyTmp(4)<>"" Then
		    Response.Write anyTmp(4)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �d���` - ���ݗ\��
        If anyTmp(5)<>"" Then
		    Response.Write anyTmp(5)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ��R�����o�ꏊ
        If anyTmp(6)<>"" Then
		    Response.Write anyTmp(6)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- 2008.01.12 NICS START -->
<% ' CY�J�b�g
        If anyTmp(14)<>"" Then
		    Response.Write anyTmp(14)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- 2008.01.12 NICS END -->

<% ' �T�C�Y
        If anyTmp(7)<>"" Then
		    Response.Write anyTmp(7)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' �^�C�v
        If anyTmp(8)<>"" Then
		    Response.Write anyTmp(8)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ����
        If anyTmp(9)<>"" Then
		    Response.Write anyTmp(9)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>

<!-- I20040223 S -->
<% ' �ގ�
        If anyTmp(12)<>"" Then
		    Response.Write anyTmp(12)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<!-- I20040223 E -->

<% ' �\��{��
        If anyTmp(10)<>"" Then
		    Response.Write anyTmp(10)
        Else
            Response.Write "<br>"
        End If
%>
                </td>
                <td nowrap align=center valign=middle>
<% ' ���o�ϖ{��
        If anyTmp(11)<>"0" Then
		    Response.Write "<a href='#' onClick='JavaScript:window.open(""bookpick.asp?line=" & LineNo &_
						   """,""pickcont"",""scrollbars=yes,resizable=yes,width=500,height=380"")'>" &_
						   anyTmp(11) & "</a>"
        Else
            Response.Write "<br>"
        End If
%>
                </td>
              </tr>
<%
    Loop
%>
<!-- �����܂� -->
            </table>
<form method=post action="bookcheck.asp">
	  <input type=hidden name="booking" value="<%=strBookingNo%>">
<% 'Mod-s 2006/03/06 h.matsuda%>
	  <INPUT type=hidden name="ShipLine" value="<%=ShipLine%>">
	  <INPUT type=hidden name="ShoriMode" value="<%=ShoriMode%>">
<%'Mod-e 2006/03/06 h.matsuda%>
      <input type=submit value="�\���f�[�^�̍X�V">
</form>
          </td>
        </tr>
      </table>
      <form action="bookcsvout.asp"><input type="submit" value="CSV�t�@�C���o��">
    �@<a href="help23.asp">CSV�t�@�C���o�͂Ƃ́H</a> 
      </form>
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
    DispMenuBarBack "bookentry.asp"
%>
</body>
</html>

