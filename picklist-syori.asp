<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
'�ʐM�����̃Z�b�g
Function SetTusinDate()
	'�߂�l		[ O ]�ʐM����(������)
				'YYYYMMDDHHNNSS
	SetTusinDate = Trim(Year(Date)) & _
				   Trim(Right("0" & Month(Date), 2)) & _
				   Trim(Right("0" & Day(Date), 2)) & _
				   Trim(Right("0" & Hour(Time), 2)) & _
				   Trim(Right("0" & Minute(Time), 2)) & _
				   Trim(Right("0" & Second(Time), 2))
End Function

' ���l���w�肵�������̕�����ɕϊ�(�E�l�E�]���ɂ͂O)�R���e�i���͗p
Function ArrangeNumV(nNumber, nFigure)
	Dim sNum

	sNum = Right(String(nFigure, "0") & CStr(nNumber), nFigure)

	ArrangeNumV = sNum
End Function

    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "index.asp"

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

	Dim bChkboxFlag
	bChkboxFlag = false

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "http://www.hits-h.com/index.asp"             '���j���[��ʂ�
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �ڍו\���s�̃f�[�^�̎擾
    Dim strData()
    LineNo=0
    Do While Not ti.AtEndOfStream
        strTemp=ti.ReadLine
        ReDim Preserve strData(LineNo)
        strData(LineNo) = strTemp
        LineNo=LineNo+1
    Loop
    ti.Close

    ' �g�����U�N�V�����t�@�C���̊g���q 
    Const SEND_EXTENT = "snd"
    ' �g�����U�N�V�����h�c
    sTranID = "EX18"
    ' �����敪
    Const sSyori = "R"
    ' ���M�ꏊ
    Const sPlace = ""

    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "expentry.asp"
    sSosin = Trim(Session.Contents("userid"))

    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
    Dim sEXxx, iSeqNo, strFileName_01, sTran, sTusin, sDate
    '�V�[�P���X�ԍ�
    iSeqNo = GetDailyTransNo
    '�ʐM�����擾
    sTusin  = SetTusinDate


	'�����ېݒ�
	If Request.Form("ok")<>"" Then

      sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
      strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
      Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

	  Set titmp=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

	  For i=1 To CInt(Request.Form("allline"))

        '�g�����U�N�V�����t�@�C���쐬
		If Request.Form("check" & i)<>"" Then
	        anyTmp=Split(strData(i-1),",")

			If anyTmp(24)<>"" Then
				sDate = Left(anyTmp(24),4) & Mid(anyTmp(24),6,2) & Mid(anyTmp(24),9,2)
			End If

	        sEXxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
	                sSosin & "," & sPlace & "," & anyTmp(25) & "," & anyTmp(4) & "," & anyTmp(3) & "," &_
					anyTmp(0) & "," & anyTmp(23) & "," & anyTmp(14) & "," & "1" & "," &_
					anyTmp(20) & "," & sDate

	        ti.WriteLine sEXxx

			bChkboxFlag = true
		End If

        ' �e���|�����t�@�C���X�V
		If Request.Form("check" & i)="" Then
            titmp.WriteLine strData(i-1)
        Else
            strTemp=anyTmp(0)
			anyTmp(26) = "1"
            For j=1 To UBound(anyTmp)
                strTemp=strTemp & "," & anyTmp(j)
            Next
            titmp.WriteLine strTemp
        End If

	  Next

      titmp.Close
      ti.Close

	  If bChkboxFlag Then
		  WriteLog fs, "a108","��R���s�b�N�A�b�v�V�X�e��-�`�^�p���ꗗ","10", ","
		  Response.Redirect "picklist.asp?kind=4"
	  End If

	Else

	  Dim sLineAry
	  For i=1 To CInt(Request.Form("allline"))
		If Request.Form("check" & i)="on" Then
			If sLineAry="" Then
				sLineAry = i
			Else
				sLineAry = sLineAry & "," & i
			End If
			bChkboxFlag = true
		End If
	  Next

	  If bChkboxFlag Then
		  Session.Contents("lines") = sLineAry
		  Response.Redirect "picklist-input.asp"
	  End If

	End If

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
<!-------------��������o�^���--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/expkoun.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strRoute = Session.Contents("route")
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
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>
            ��R���s�b�N�A�b�v���ꗗ�i�`�^�p�j</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
	 	<BR>

<%
	strError = "�m�F�^�ύX�`�F�b�N�{�b�N�X�̃`�F�b�N���t���Ă��܂���B"
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError 
%>
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
<!-------------�o�^��ʏI���--------------------------->
<%
    strTemp = "picklist.asp?kind=4"
    DispMenuBarBack strTemp
%>
</body>
</html>
<%

%>
