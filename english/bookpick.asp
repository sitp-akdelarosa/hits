<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "EXPORT", "expentry.asp"

	Dim strBookingNo
	strBookingNo = Request.QueryString("line")

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

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' �\���t�@�C���̃��R�[�h������ԌJ��Ԃ�
    LineNo=0
    Do While Not ti.AtEndOfStream
        LineNo=LineNo+1
		If LineNo=CInt(strBookingNo) Then
			anyTmp=Split(ti.ReadLine,",")
		Else
			strDam = ti.ReadLine
		End If
	Loop

	ti.close()

    ' �A�o�R���e�i�Ɖ�X�g�\��
'    WriteLog fs, "1011","�u�b�L���O���Ɖ�-���o�σR���e�i���","00", anyTmp(1) & "/" & anyTmp(12) & ","	'D20040223
    WriteLog fs, "1011","�u�b�L���O���Ɖ�-���o�σR���e�i���","00", anyTmp(1) & "/" & anyTmp(13) & ","	'I20040223

%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="JavaScript:this.focus()">
<!-------------��������ꗗ���--------------------------->
      <center>
      <table>
        <tr>
          <td align=center> 
            <table>
              <tr>
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�s�b�N�A�b�v�σR���e�i���</b></td>
              </tr>
            </table>
            <br>

<%
	If anyTmp(11)<>"0" Then
%>
            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center"> 
                <td nowrap bgcolor="#000099"><font color="#ffffff"><b>Booking No.</td>
				</td>
                <td nowrap bgcolor="#ffffff"><%=anyTmp(1)%></td>
				</td>
			  </tr>
			</table>
			<BR>

            <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
              <tr align="center" bgcolor="#FFCC33"> 
                <td nowrap>�R���e�iNo.</td>
                <td nowrap>TW(Kg)</td>
              </tr>
<!-- ��������f�[�^�J��Ԃ� -->
<%
'		contTmp=Split(anyTmp(12),"/")	'D20040223
		contTmp=Split(anyTmp(13),"/")	'I20040223
		MaxCont=UBound(contTmp)
		For k=0 To MaxCont
			lineTmp=Split(contTmp(k),"!")			' Ins 2005/03/28
%>
              <tr bgcolor="#FFFFFF"> 
				<td nowrap align=center valign=middle>
<% ' �R���e�iNo.
'	        If contTmp(k)<>"" Then								' Del 2005/03/28
'			    Response.Write contTmp(k)						' Del 2005/03/28
	        If lineTmp(0)<>"" Then								' Ins 2005/03/28
			    Response.Write lineTmp(0)						' Ins 2005/03/28
	        Else
	            Response.Write "<br>"
	        End If
%>
                </td>
				<td nowrap align=center valign=middle>
<% ' TW															' Ins 2005/03/28
	        If lineTmp(1)<>"" Then								' Ins 2005/03/28
			If lineTmp(1)="0" Then
			    Response.Write "�|"
			ElseIf Len(lineTmp(1))<=2 Then
			    Response.Write lineTmp(1) & "00"
			Else
			    Response.Write lineTmp(1)
			End If
	        Else												' Ins 2005/03/28
	            Response.Write "�|"
	        End If												' Ins 2005/03/28
%>
                </td>
              </tr>
<%
    	Next
%>
<!-- �����܂� -->
            </table>
<% End If %>

      <form>
		<input type=button value="Close" onClick="JavaScript:window.close()">
      </form>
      </center>
    </td>
  </tr>
</table>

</body>
</html>

