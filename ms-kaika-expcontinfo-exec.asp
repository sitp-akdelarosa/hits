<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	�y�R���e�i�����́z	�G���[�`�F�b�N�A�\���A�t�@�C���쐬
	'
%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"

	Dim bKind,iLineNo
	' �V�K(1) or �X�V(0)
    bKind = Trim(Request.form("kind"))
    iLineNo	= CInt(Trim(Request.form("lineno")))

	' �g�����U�N�V�����t�@�C���̊g���q 
	Const SEND_EXTENT = "snd"
	' �g�����U�N�V�����h�c
	Const sTranID05 = "EX05"
	Const sTranID16 = "EX16"
	' �����敪
	Const sSyori = "R"
	' ���M�ꏊ
	Const sPlace = ""
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "ms-kaika.asp"
	sSosin = Trim(Session.Contents("userid"))
    ' �G���[�t���O�̃N���A
    bError = false
    ' ���̓t���O�̃N���A
    bInput = true
    ' �w������̎擾
	Dim sUser,sUserNo,sVslCode,sVoyCtrl,sBooking,sCont,sSeal,sCargoWeight,sContWeight,sRifer,sDanger
    sUser 	= UCase(Trim(Request.form("user")))
    sUserNo = UCase(Trim(Request.form("userno")))
    sVslCode = UCase(Trim(Request.form("vslcode")))
    sVoyCtrl = UCase(Trim(Request.form("voyctrl")))
    sBooking = UCase(Trim(Request.form("booking")))
	sCont 		= UCase(Trim(Request.form("cont")))
	sSeal 		= UCase(Trim(Request.form("seal")))
	sCargoWeight= UCase(Trim(Request.form("cargow")))
    If sCargoWeight<>"" Then
        sCargoWeight=CInt(CDbl(sCargoWeight)*10)
    End If
	sContWeight	= UCase(Trim(Request.form("contw")))
    If sContWeight<>"" Then
        sContWeight=CInt(CDbl(sContWeight)*10)
    End If
	sRifer 		= UCase(Trim(Request.form("rifer")))
	If sRifer="ON" Then
		sRifer = "R"
	Else
		sRifer = ""
	End If
	sDanger 	= UCase(Trim(Request.form("danger")))
	If sDanger="ON" Then
		sDanger = "H"
	Else
		sDanger = ""
	End If
'	iLineNo		= Request.form("lineno")

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' ���p�J���}�`�F�b�N
	If InStr(sCont,",")<>0 Or _
		InStr(sSeal,",")<>0 _
	Then

	    bError = true
		strError = "���͂̍ہA���p�J���}�͎g�p���Ȃ��ŉ������B"

	Else

	    ConnectSvr conn, rsd
		' �R���e�iNo.�����݂��邩
		sql = "SELECT ExportCont.VanTime,VslSchedule.ShipLine " & _
			  "FROM ExportCont,VslSchedule " & _
			  "WHERE " & _
				"ExportCont.VslCode='" & sVslCode & "' AND " & _
				"ExportCont.ContNo='" & sCont & "' AND " & _
				"ExportCont.BookNo='" & sBooking & "' AND " & _
				"VslSchedule.VslCode='" & sVslCode & "'"
		rsd.Open sql, conn, 0, 1, 1

		Dim sVanTime,sShipLine
		If Not rsd.EOF Then
		    sVanTime  = Trim(rsd("VanTime"))
		    sShipLine = Trim(rsd("ShipLine"))
		Else
''		    bError = true
''			strError = "�w�肳�ꂽ�R���e�iNo.�����݂��܂���B"
		End If
		rsd.Close

	End If

    If Not bError Then

' �g�����U�N�V�����t�@�C���쐬
	    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
	    Dim sEX16, sEX05, iSeqNo_EX16, iSeqNo_EX05, strFileName, sTran, sTusin, sDate

		'�ʐM�����擾
		sTusin  = SetTusinDate

		If bKind=1 Then
			' EX16�p�p�����[�^�̎擾
			Dim sContSize,sContType,sContHeight,sRemark,sTrucker,sWHArTime,sCYRecDate,sPickPlace
			sql = "SELECT ContSize,ContType,ContHeight,Remark,Trucker,WHArTime,CYRecDate,PickPlace " & _
				  "FROM ExportCargoInfo " & _
				  "WHERE " & _
					"Shipper='" & sUser & "' AND " & _
					"ShipCtrl='" & sUserNo & "'"
			rsd.Open sql, conn, 0, 1, 1

			If Not rsd.EOF Then
			    sContSize 	= Trim(rsd("ContSize"))
			    sContType 	= Trim(rsd("ContType"))
			    sContHeight = Trim(rsd("ContHeight"))
			    sRemark 	= Trim(rsd("Remark"))
			    sTrucker 	= Trim(rsd("Trucker"))
			    sWHArTime 	= Trim(rsd("WHArTime"))
			    sCYRecDate 	= Trim(rsd("CYRecDate"))
			    sPickPlace 	= Trim(rsd("PickPlace"))
			Else
			    bError = true
				strError = "�w�肳�ꂽ�R���e�iNo.�����݂��܂���B"
			End If
			rsd.Close

			'�V�[�P���X�ԍ�
			iSeqNo_EX16 = GetDailyTransNo
			'�q�ɓ����w�����
			If sWHArTime<>"" And Not IsNull(sWHArTime) Then
				sWHArTime = "20" & Left(sWHArTime,2) & Mid(sWHArTime,4,2) & Mid(sWHArTime,7,2) & _
							Mid(sWHArTime,10,2) & Mid(sWHArTime,13,2)
			End If
			'�b�x�����w���
			If sCYRecDate<>"" And Not IsNull(sCYRecDate) Then
				sCYRecDate = "20" & Left(sCYRecDate,2) & Mid(sCYRecDate,4,2) & Mid(sCYRecDate,7,2)
			End If

			sEX16 = iSeqNo_EX16 & "," & sTranID16 & "," & sSyori & ","  & sTusin & ",Web - " & _
					sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
					sBooking & "," & sUser & "," & sUserNo & "," & sSosin & "," & _
					sCont & "," & sContSize & "," & sContType & "," & sContHeight & "," & _
					sRemark & "," & sTrucker & "," & _
					sWHArTime & "," & sCYRecDate & "," & sPickPlace
			sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX16
			strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
		    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
			ti.WriteLine sEX16
		    ti.Close
			Set ti = Nothing

		End If

		'�V�[�P���X�ԍ�
		iSeqNo_EX05 = GetDailyTransNo
		'�o���l�ߓ���
		If sVanTime<>"" Then
			sVanTime = "20" & Left(sVanTime,2) & Mid(sVanTime,4,2) & Mid(sVanTime,7,2) & _
						Mid(sVanTime,10,2) & Mid(sVanTime,13,2)
		End If

		sEX05 = iSeqNo_EX05 & "," & sTranID05 & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sCont & "," & sBooking & "," & sShipLine & "," & sVanTime & "," & _
				sContWeight & "," & sSeal & "," & sCargoWeight & "," & sSosin & ",," & _
				sRifer & sDanger
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX05
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sEX05
	    ti.Close
		Set ti = Nothing

' �g�����U�N�V���������܂�


' Temp�t�@�C���쐬

	    ' File System Object �̐���
'	    Set fs=Server.CreateObject("Scripting.FileSystemobject")

	    Dim strTempFileName
	    ' �\���t�@�C���̎擾
	    strTempFileName = Session.Contents("tempfile")
	    If strTempFileName="" Then
	        ' �Z�b�V�������؂�Ă���Ƃ�
	        Response.Redirect "http://www.hits-h.com/index.asp"             '���j���[��ʂ�
	        Response.End
	    End If

	    strTempFileName="./temp/" & strTempFileName

	    ' �\���t�@�C����Open
	    Set ti=fs.OpenTextFile(Server.MapPath(strTempFileName),1,True)

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

	    Set ti=fs.OpenTextFile(Server.MapPath(strTempFileName),2,True)

  		anyTmp=Split(strData(iLineNo-1),",")
        anyTmp(0) = sVslCode
        anyTmp(1) = sVoyCtrl
        anyTmp(2) = sUser
        anyTmp(3) = sUserNo
        anyTmp(4) = sBooking
        anyTmp(5) = sCont
        anyTmp(6) = sSeal
        If sCargoWeight<>"" Then
            anyTmp(7) = sCargoWeight/10
        End If
        If sContWeight<>"" Then
            anyTmp(8) = sContWeight/10
        End If
        anyTmp(9) = sRifer
        anyTmp(10) = sDanger

        For i=1 To LineNo
            If i<>CInt(iLineNo) Then
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

' Temp�����܂�

	End If

	' ���O�t�@�C�������o��
    If sCargoWeight<>"" Then
        sCargoWeight = sCargoWeight/10
    End If
    If sContWeight<>"" Then
        sContWeight = sContWeight/10
    End If
	strOption = sCont & _
				"/" & sSeal & _
				"/" & sCargoWeight &_
				"/" & sContWeight & _
				"/" & sRifer & _
				"/" & sDanger & ","
    If bError Then
		strOption = strOption &	"���͓��e�̐���:1(���)"
    Else
		strOption = strOption & "���͓��e�̐���:0(������)"
    End If

    WriteLog fs, "4106","�C�ݓ��͗A�o�R���e�i���-������","10", strOption


	If Not bError Then
		Response.Redirect "ms-kaika-expcontinfo-list.asp"
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
<!-------------�������烍�O�C�����͉��--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika5t.gif" width="506" height="73"></td>
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
                <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                <td nowrap><b>�X�V������</b></td>
                <td><img src="gif/hr.gif"></td>
              </tr>
            </table>
            <br>
<%
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
<!-------------���O�C����ʏI���--------------------------->
<%
    DispMenuBarBack "JavaScript:window.history.back()"
%>
</body>
</html>

<%
%>
