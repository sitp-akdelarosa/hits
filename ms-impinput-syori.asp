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
    CheckTempFile "MSIMPORT", "impentry.asp"

    ' �w������̎擾
    Dim strKind       '���͎��(1=�͎���,2=��������)
    Dim iLine         '���͍s
    Dim strRequest    '�߂��
    strKind=Session.Contents("editkind")
    iLine=CInt(Session.Contents("editline"))
    strRequest=Session.Contents("request")

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
    If strKind="1" Then
        sTranID = "IM10"
    Else
        sTranID = "IM11"
    End If
    ' �����敪
    Const sSyori = "R"
    ' ���M�ꏊ
    Const sPlace = ""

    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "ms-impentry.asp&kind=2"
    sSosin = Trim(Session.Contents("userid"))

	' �^�C�g���擾
	strTitle = Trim(Request.form("title"))

    ' �G���[�t���O�̃N���A
    bError = false

    If Not bError Then
        '�g�����U�N�V�����t�@�C���쐬
        anyTmp=Split(strData(iLine-1),",")
        ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
        Dim sIMxx, iSeqNo, strFileName_01, sTran, sTusin, sDate
        '�V�[�P���X�ԍ�
        iSeqNo = GetDailyTransNo
        '�ʐM�����擾
        sTusin  = SetTusinDate
        sDate = Trim(Request.form("Year")) 
        sDate = sDate & Right("0" & Trim(Request.form("Month")),2)
        sDate = sDate & Right("0" & Trim(Request.form("Day")),2)
        sDate = sDate & Right("0" & Trim(Request.form("Hour")),2)
        sDate = sDate & Right("0" & Trim(Request.form("Min")),2)

        If strKind="1" Then
            sIMxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
                    sSosin & "," & sPlace & "," & anyTmp(4) & "," &  anyTmp(3) & "," & _
                    anyTmp(1) & "," & anyTmp(0) & "," & sDate & "," & sSosin
        Else
            sIMxx = iSeqNo & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
                    sSosin & "," & sPlace & "," & anyTmp(4) & "," &  anyTmp(3) & "," & _
                    anyTmp(1) & "," & anyTmp(0) & "," & sDate & "," & sSosin
        End If
        sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo
        strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
        ti.WriteLine sIMxx
        ti.Close

        sLogDate = Trim(Request.form("Year")) & "/"
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Month")),2) & "/"
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Day")),2) & " "
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Hour")),2) & ":"
        sLogDate = sLogDate & Right("0" & Trim(Request.form("Min")),2)
        strOption = anyTmp(0) & "/" & sLogDate & "," & "���͓��e�̐���:0(������)"

        ' �e���|�����t�@�C���X�V
        strTemp=Left(sDate,4) & "/" & Mid(sDate,5,2) & "/" & Mid(sDate,7,2) & " " & Mid(sDate,9,2) & ":" & Mid(sDate,11,2)
        If strKind="1" Then
            anyTmp(44)=strTemp
        Else
            anyTmp(45)=strTemp
        End If
        Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
        For i=1 To LineNo
            If i<>CInt(iLine) Then
                ti.WriteLine strData(i-1)
            Else
                strTemp=anyTmp(0)
                For j=1 To UBound(anyTmp)
                    strTemp=strTemp & "," & anyTmp(j)
                Next
                ti.WriteLine strTemp
            End If
        Next
        ti.Close

        ' �C�ݓ��͍��ڑI��
        If strKind="1" Then
            WriteLog fs, "","�A�o���Ɩ��x��-�A��������R���e�i�q�ɓ�����������","10", strOption
        Else
            WriteLog fs, "2107","�A���R���e�i�Ɖ�-�f�o���j���O������������","10", strOption
        End If

        If strRequest="ms-impdetail.asp" Then
            strTemp=strRequest & "?line=" & iLine
        Else
            strTemp=strRequest
        End If

        ' �߂��ʂփ��_�C���N�g
        Response.Redirect strTemp
        Response.End
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
          <td rowspan=2><img src="gif/imprikuun.gif" width="506" height="73"></td>
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
				<%=strRoute%> &gt; ��������
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
    strOption = anyTmp(1) & "/" & sLogDate & "," & "���͓��e�̐���:1(���)"
    ' �C�ݓ��͍��ڑI��
    If strKind="1" Then
        WriteLog fs, "","�A�o���Ɩ��x��-�A��������R���e�i�q�ɓ�����������","10", strOption
    Else
        WriteLog fs, "2107","�A���R���e�i�Ɖ�-�f�o���j���O������������","10", strOption
    End If
%>
      <br><br>
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
    If strRequest="ms-impdetail.asp" Then
        strTemp=strRequest & "?line=" & iLine
    Else
        strTemp=strRequest
    End If
    DispMenuBarBack strTemp
%>
</body>
</html>
<%

%>