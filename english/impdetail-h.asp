<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "IMPORT", "impentry.asp"

    ' �w������̎擾
    Dim iLineNo
    iLineNo = Request.QueryString("line")

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

    ' �ڍו\���s�̃f�[�^�̎擾
    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
        If iLineNo=LineNo Then
           Exit Do
        End If
    Loop
    ti.Close

    ' �A���R���e�i�Ɖ�
    WriteLog fs, "2007","�A���R���e�i�Ɖ�-�P�ƃR���e�i�ېŗA������","00", anyTmp(1) & ","
%>

<html>
<head>
<title>�ېŗA������</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body bgcolor="#E6E8FF" text="#000000">
  <center>
<!-----�ېŗA������---------------->
  <table>
    <tr>
      <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
      <td nowrap><b>�ېŗA������</b></td>
    </tr>
  </table>
  <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
    <tr align="center" bgcolor="#FFCC33"> 
      <td nowrap>FROM</td>
      <td nowrap>TO</td>
    </tr>
    <tr align="center"> 
      <td align=center>
<% ' �ېŗA��(From)
    Response.Write DispDateTimeCell(anyTmp(28),5)
%>
      </td>
      <td align=center>
<% ' �ېŗA��(To)
    Response.Write DispDateTimeCell(anyTmp(29),5)
%>
      </td>
    </tr>
  </table>
  <FORM>
    <INPUT type="button" value=" Close " onClick="opener.winfl=0;window.close()">
  </FORM>
  </center>
</body>
</html>
