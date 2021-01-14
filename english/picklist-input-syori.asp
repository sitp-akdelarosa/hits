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

    ' �w������̎擾
    Dim iLine         '���͍s
    iLine=Session.Contents("lineary")
	Session.Contents("lineary") = ""

	sLoginKind = Session.Contents("userkind")

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

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

	' �^�C�g���擾
	strTitle = Trim(Request.form("title"))


    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
    Dim sEXxx, iSeqNo, strFileName_01, sTran, sTusin, sDate, sPickPlace

    '�V�[�P���X�ԍ�
    iSeqNo = GetDailyTransNo

    '�ʐM�����擾
    sTusin  = SetTusinDate
    sPickPlace = Trim(Request.form("pickplace")) 
    sDate = Trim(Request.form("pickyear")) 
    sDate = sDate & Right("0" & Trim(Request.form("pickmon")),2)
    sDate = sDate & Right("0" & Trim(Request.form("pickday")),2)
	If sDate="00" Then
		sDate = ""
	End If
    strTemp=Left(sDate,4) & "/" & Mid(sDate,5,2) & "/" & Mid(sDate,7,2)

    sLogDate = Trim(Request.form("pickyear")) & "/"
    sLogDate = sLogDate & Right("0" & Trim(Request.form("pickmon")),2) & "/"
    sLogDate = sLogDate & Right("0" & Trim(Request.form("pickday")),2)
	If sLogDate="/0/0" Then
		sLogDate=""
	End If

	If sLoginKind="�`�^" Then
	    strOption = sPickPlace & "/" & sLogDate & "," & "���͓��e�̐���:0(������)"
	Else
	    strOption = sLogDate & "," & "���͓��e�̐���:0(������)"
	End If

    sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
    strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)

	iLineAry = Split(iLine,",")
    '�g�����U�N�V�����t�@�C���쐬
	For k=0 To UBound(iLineAry)
	    ' �G���[�t���O�̃N���A
	    bError = false

	    If Not bError Then
		    '�V�[�P���X�ԍ�
		    iSeqNo = GetDailyTransNo

	        anyTmp=Split(strData(iLineAry(k)-1),",")
	        sEXxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
	                sSosin & "," & sPlace & "," & anyTmp(25) & "," & anyTmp(4) & "," & anyTmp(3) & "," &_
					anyTmp(0) & "," & anyTmp(23) & "," & anyTmp(14) & ",2," &_
					sPickPlace & "," & sDate

	        ti.WriteLine sEXxx
	    End If
	Next

    ti.Close

    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)

    ' �e���|�����t�@�C���X�V
    For i=1 To LineNo
		bSameFlag = false
		For l=0 To UBound(iLineAry)
			If i=CInt(iLineAry(l)) Then
				bSameFlag = true
			End If
		Next

        If Not bSameFlag Then
            ti.WriteLine strData(i-1)
        Else
	        anyTmp=Split(strData(i-1),",")
            strTemp=anyTmp(0)
			If sPickPlace<>"" Then
				anyTmp(20) = sPickPlace
				anyTmp(27) = "1"
				anyTmp(26) = "2"
			End If
			If sLogDate<>"" Then
			    anyTmp(24) = sLogDate
				anyTmp(28) = "1"
				anyTmp(26) = "2"
			End If
            For j=1 To UBound(anyTmp)
                strTemp=strTemp & "," & anyTmp(j)
            Next
            ti.WriteLine strTemp
        End If
    Next

    ti.Close

    ' �C�ݓ��͍��ڑI��
	If sLoginKind="�`�^" Then
	    WriteLog fs, "a109","��R���s�b�N�A�b�v�V�X�e��-��R�����ꏊ�E���o������","12", strOption
	    Response.Redirect "picklist.asp?kind=4"
	Else
	    WriteLog fs, "a109","��R���s�b�N�A�b�v�V�X�e��-��R�����ꏊ�E���o������","11", strOption
		Response.Redirect "picklist.asp?kind=2"
	End If

    Response.End

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
				<%=strRoute%> &gt; ��R�����ꏊ�E���o������
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
<%
    Response.Write strTitle
%>
            ����</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
	 	<BR>

<%
    ' �G���[���b�Z�[�W�̕\��
    DispErrorMessage strError 
    strOption = sPickPlace & "/" & sLogDate & "," & "���͓��e�̐���:1(���)"
    ' �C�ݓ��͍��ڑI��
    WriteLog fs, "a109","��R���s�b�N�A�b�v�V�X�e��-��R�����ꏊ�E���o������","10", strOption
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
    strTemp = "picklist.asp?kind=" & iLoginKind
    DispMenuBarBack strTemp
%>
</body>
</html>
<%

%>