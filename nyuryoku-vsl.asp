<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="vessel.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-in1.asp"

	' �g�����U�N�V�����t�@�C���̊g���q 
	Const SEND_EXTENT = "snd"

	' ���M��
	sSosin = Trim(Session.Contents("userid"))

	' ���M�ꏊ
	Const sPlace = ""

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName,strFileNameP
    strFileName = Session.Contents("tempfile")

    If strFileName="" Then
        ' �����w��̂Ȃ��Ƃ�
        strFileName="test.csv"
    End If
    strFileNameP="./temp/" & strFileName

    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileNameP),1,True)

    ' �{�D���Ê�{���̎擾
    If Not ti.AtEndOfStream Then
        anyTmp=Split(ti.ReadLine,",")
    End If

'�u�r�O�P�Z�b�g
	Dim sVs01, iSeqNo_VS01, sTran1, sSyori, sTusin1, sCall, sAge, sTumi, sOpe

	'�V�[�P���X�ԍ�
	iSeqNo_VS01 = GetDailyTransNo

	'�g�����U�N�V�����h�c
	sTran1 = "VS01"

	'�����敪
	sSyori = "R"

	'�ʐM�����擾
	sTusin  = SetTusinDate

    ' �D���̕\��(�R�[���T�C��)
    sCall = anyTmp(2)

    ' �g�����q�̕\��
    sAge  = anyTmp(5)

	' �ςݎ��q�̕\��
	sTumi = anyTmp(6)

    ' �^�s�D�ЃR�[�h
	sShipLine = anyTmp(0)

	sVs01 = iSeqNo_VS01 & "," & sTran1 & "," & sSyori & ","  & sTusin & ",Web - " & _
			sSosin & "," & sPlace & "," & sCall & "," &  sAge & "," & sTumi & "," &  sShipLine

    ' �ڍו\���s�̃f�[�^���̎擾
    If Not ti.AtEndOfStream Then
        iCount=CInt(ti.ReadLine)
    End If

'�u�r�O�Q�Z�b�g
	'VS02�w�b�_�[���Z�b�g

	'�g�����U�N�V�����h�c
	sTran1 = "VS02"

	sVs02 = ""
	sVs02 = iSeqNo_VS01 & "," & sTran1 & "," & sSyori & ","  & sTusin & ",Web - " & _
			sSosin & "," & sPlace & "," & sCall & "," &  sAge & "," & sTumi

    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
		sVs02 = sVs02 & "," & anyTmp(0) & "," & Left(SetTusinDate2(anyTmp(6)),8) & "," & _
		        Left(SetTusinDate2(anyTmp(7)),8) & "," & SetTusinDate2(anyTmp(2)) & "," & _
				SetTusinDate2(anyTmp(3)) & ",," & SetTusinDate2(anyTmp(5))
    Loop
    ti.Close

'�g�����U�N�V�����f�[�^�o�^
	'�t�@�C�����̂̎擾
	Dim sFileName
	sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2)
	strFileName_01 = "./send/" & sFileName & iSeqNo_VS01 & "." & SEND_EXTENT
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
	ti.WriteLine sVs01
	ti.WriteLine sVs02
    ti.Close

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
          </td></tr>
      </table>
<%
		strError = "����ɍX�V����܂����B"
	    DispInformationMessage strError
%>
      <form action="nyuryoku-in1.asp">
        <br>�@<br>
        <input type=submit value="   ��  ��   ">
      </form>

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
<!-------------�ꗗ��ʏI���--------------------------->
<%
    DispMenuBarBack "nyuryoku-in1.asp"
%>
</body>

<%
    ' �{�D���Ó��͈ꗗ
    WriteLog fs, "�{�D���Ó��͈ꗗ", "�o�^����"
%>
