<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	�y�A���R���e�i�����́z	�G���[�`�F�b�N�A�\���A�t�@�C���쐬
	'
%>

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-kaika.asp"

	Dim bKind,sSend,sStop,sDel,iLineNo
	' �V�K(1) or �X�V(0)
    bKind = Trim(Session.Contents("kind"))
	' ���
    sSend 	= Trim(Request.form("send"))
    sStop 	= Trim(Request.form("stop"))
    sDel 	= Trim(Request.form("del"))
    iLineNo	= Trim(Request.form("lineno"))

	If bKind=1 And sStop<>"" Then
        Response.Redirect "ms-kaika-impcontinfo-updatecheck.asp"

	ElseIf bKind=2 And sStop<>"" Then
        Response.Redirect "ms-kaika-impcontinfo-list.asp"

	Else
		' �g�����U�N�V�����t�@�C���̊g���q 
		Const SEND_EXTENT = "snd"
		' �g�����U�N�V�����h�c
		Const sTranID = "IM18"
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
	    Dim sVslCode,sVoyCtrl,sUser,sCont,sBL,sTraderCode,iSize,sType,sRemark,sEmpDate,sEmpDateS
	    sUser 	= UCase(Trim(Request.form("user")))
	    sCont 	= UCase(Trim(Request.form("cont")))
	    sVslCode = UCase(Trim(Request.form("vslcode")))
	    sVoyCtrl = UCase(Trim(Request.form("voyctrl")))
	    sBL		 = UCase(Trim(Request.form("bl")))
	    sTraderCode = UCase(Trim(Request.form("tradercode")))
	    iSize 	= Trim(Request.form("size"))
	    sType 	= UCase(Trim(Request.form("type")))
	    sRemark = UCase(Trim(Request.form("remark")))
		sEmpDate = Trim(Request.form("emparvtime_year")) 
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_mon")),2)
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_day")),2)
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_hour")),2)
		sEmpDate = sEmpDate & Right("0" & Trim(Request.form("emparvtime_min")),2)
		If sEmpDate=0000 Then
			sEmpDate = ""
		Else
			sEmpDateT = Trim(Request.form("emparvtime_year")) 
			sEmpDateT = sEmpDateT & "/" & Right("0" & Trim(Request.form("emparvtime_mon")),2)
			sEmpDateT = sEmpDateT & "/" & Right("0" & Trim(Request.form("emparvtime_day")),2)
			sEmpDateT = sEmpDateT & " " & Right("0" & Trim(Request.form("emparvtime_hour")),2)
			sEmpDateT = sEmpDateT & ":" & Right("0" & Trim(Request.form("emparvtime_min")),2)
			sEmpDateS = Trim(Request.form("emparvtime_year")) 
			sEmpDateS = sEmpDateS & "�N " & Trim(Request.form("emparvtime_mon"))
			sEmpDateS = sEmpDateS & "�� " & Trim(Request.form("emparvtime_day"))
			sEmpDateS = sEmpDateS & "�� " & Trim(Request.form("emparvtime_hour"))
			sEmpDateS = sEmpDateS & "�� " & Trim(Request.form("emparvtime_min")) & "�� "
		End If


	    ' File System Object �̐���
	    Set fs=Server.CreateObject("Scripting.FileSystemobject")

		' ���p�J���}�`�F�b�N
		If InStr(sVslCode,",")<>0 Or _
			InStr(sVoyCtrl,",")<>0 Or _
			InStr(sBL,",")<>0 Or _
			InStr(sTraderCode,",")<>0 Or _
			InStr(sRemark,",")<>0 Or _
			InStr(sCont,",")<>0 Or _
			InStr(sUser,",")<>0 _
		Then

		    bError = true
			strError = "���͂̍ہA���p�J���}�͎g�p���Ȃ��ŉ������B"

		Else

			ConnectSvr conn, rsd
			' �D���Ǝ��q�ƃR���e�iNo.�̏d���`�F�b�N
			Dim iRecCount
			If Not bKind=0 Then
				sql = "SELECT count(*) FROM ImportCargoInfo " & _
						"WHERE VslCode='" & sVslCode & "' AND DsVoyage='" & sVoyCtrl & "' AND ContNo='" & sCont & "'"
				rsd.Open sql, conn, 0, 1, 1
				If Not rsd.EOF Then
					iRecCount = rsd(0)
					If Not iRecCount=0 Then
					    bError = true
						strError = "�D��, Voyage No, �R���e�iNo.���d�����Ă��܂��B"
					End If
				End If
				rsd.Close
			End If

			' �D�������݂��邩
			sql = "SELECT count(*) FROM mVessel WHERE VslCode='" & sVslCode & "'"
			rsd.Open sql, conn, 0, 1, 1
			If Not rsd.EOF Then
				iRecCount = rsd(0)
				If iRecCount=0 Then
				    bError = true
					strError = "�w�肳�ꂽ�D�������݂��܂���B"
				End If
			End If
			rsd.Close

		End If
	End If

    If Not bError Then
		' �����敪
		Dim sSyori
		If sSend<>"" Then
			sSyori = "R"
		Else
			sSyori = "D"
		End If

		Const sContainer = ""

' �g�����U�N�V�����t�@�C���쐬

	    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
	    Dim sIM18, iSeqNo_IM18, strFileName, sTran, sTusin, sDate
		'�V�[�P���X�ԍ�
		iSeqNo_IM18 = GetDailyTransNo
		'�ʐM�����擾
		sTusin  = SetTusinDate

		sIM18 = iSeqNo_IM18 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sBL & "," & sUser & "," &  sSosin & "," & _
				sCont & "," & iSize & "," & sType & "," & sTraderCode & "," & _
				sRemark & "," & sEmpDate
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_IM18
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sIM18
	    ti.Close
		Set ti = Nothing

' �g�����U�N�V���������܂�


' Temp�t�@�C���쐬

		    ' File System Object �̐���
		    Set fs=Server.CreateObject("Scripting.FileSystemobject")

		    Dim strTempFileName
			If bKind=1 Then
			    ' �e���|�����t�@�C�������쐬���āA�Z�b�V�����ϐ��ɐݒ�
			    strTempFileName = GetNumStr(Session.SessionID, 8) & ".csv"
			    Session.Contents("tempfile")=strTempFileName

			Else
			    ' �\���t�@�C���̎擾
			    strTempFileName = Session.Contents("tempfile")
			    If strTempFileName="" Then
			        ' �Z�b�V�������؂�Ă���Ƃ�
			        Response.Redirect "http://www.hits-h.com/index.asp"             '���j���[��ʂ�
			        Response.End
			    End If

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

		' �X�V��
			If bKind=0 Then

	      		anyTmp=Split(strData(iLineNo-1),",")
	            anyTmp(0) = sVslCode
	            anyTmp(1) = sVoyCtrl
	            anyTmp(2) = sUser
	            anyTmp(3) = sBL
	            anyTmp(4) = sCont
	            anyTmp(5) = sTraderCode
	            anyTmp(6) = sEmpDateT
	            anyTmp(7) = iSize
	            anyTmp(8) = sType
	            anyTmp(9) = sRemark

	            For i=1 To LineNo
	                If i<>CInt(iLineNo) Then
	                    ti.WriteLine strData(i-1)
	                Else
						If sDel="" Then
		                    strTemp=anyTmp(0)
		                    For j=1 To UBound(anyTmp)
		                        strTemp=strTemp & "," & anyTmp(j)
		                    Next
		                    ti.WriteLine strTemp
						End If
	                End If
	            Next
	            ti.Close

		' �V�K�o�^��
			Else

				Dim strTemp

				If bKind=2 Then
		            For i=1 To LineNo
						ti.WriteLine strData(i-1)
					Next
				End If

				strTemp = sVslCode & "," &  sVoyCtrl & "," & sUser & "," & sBL & "," & _
						 sCont & "," & sTraderCode & "," & sEmpDateT & "," & _
						 iSize & "," & sType & "," & sRemark

                ti.WriteLine strTemp
	            ti.Close

			End If

		End If

' Temp�����܂�

	' ���O�t�@�C�������o��
	Dim sLogTime
	sLogTime = Trim(Request.form("emparvtime_year")) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_mon")),2) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_day")),2) & " "
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_hour")),2) & ":"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_min")),2)
	If sEmpDateT="" Then
		sLogTime = ""
	End If

	strOption = sVslCode & _
				"/" & sVoyCtrl & _
				"/" & sUser & _
				"/" & sBL & _
				"/" & sCont & _
				"/" & sTraderCode & _
				"/" & sLogTime & _
				"/" & iSize & _
				"/" & sType & _
				"/" & sRemark & ","

    If bError Then
		strOption = strOption &	"���͓��e�̐���:1(���)"
    Else
		strOption = strOption & "���͓��e�̐���:0(������)"
    End If

	If bKind=1 Then
  		WriteLog fs, "4110","�C�ݓ��͗A���R���e�i���-������","11", strOption
	ElseIf sDel<>"" Then
  		WriteLog fs, "4110","�C�ݓ��͗A���R���e�i���-������","13", strOption
	Else
  		WriteLog fs, "4110","�C�ݓ��͗A���R���e�i���-������","12", strOption
	End If


    If Not bError And bKind=0 Then
		Response.Redirect "ms-kaika-impcontinfo-list.asp"
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
          <td rowspan=2><img src="gif/kaika6t.gif" width="506" height="73"></td>
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
          <td><img src="gif/botan.gif" width="17" height="17"></td>
<% If bKind<>0 Then %>
          <td nowrap><b>�V�K������</b></td>
<% Else %>
          <td nowrap><b>�X�V������</b></td>
<% End If %>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
<%
	If Not bError Then
%>

              <table border="1" cellspacing="2" cellpadding="3">

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�D��(�R�[���T�C��)</b></font>
                  </td>
                  <td nowrap>
					<%=sVslCode%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Voyage No.</b></font>
                  </td>
                  <td nowrap>
					<%=sVoyCtrl%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�C�݃R�[�h</b></font>
                  </td>
                  <td nowrap>
					<%=sSosin%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle> <font color="#FFFFFF"><b>�׎�R�[�h</b></font></td>
                  <td nowrap>
					<%=sUser%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>BL No.</b></font>
                  </td>
                  <td nowrap>
					<%=sBL%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�R���e�iNo.</b></font>
                  </td>
                  <td nowrap>
					<%=sCont%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�w�藤�^�Ǝ҃R�[�h</b></font>
                  </td>
                  <td nowrap>
					<% If sTraderCode<>"" Then %>
						<%=sTraderCode%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>������q�ɓ����w�����</b></font>
                  </td>
                  <td nowrap>
					<% If sEmpDate<>"" Then %>
						<%=sEmpDateS%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�T�C�Y</b></font>
                  </td>
                  <td nowrap>
					<% If iSize<>"" Then %>
						<%=iSize%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�^�C�v</b></font>
                  </td>
                  <td nowrap>
					<% If sType<>"" Then %>
						<%=sType%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�q�ɗ���</b></font>
                  </td>
                  <td nowrap>
					<% If sRemark<>"" Then %>
						<%=sRemark%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

              </table><BR>
<%
	    ' �G���[���b�Z�[�W�̕\��
		strError = "����ɑ��M����܂����B"

		Session.Contents("kind") = 2

	End If

		If bError Then
%><BR><%
	        DispErrorMessage strError
		Else
	        DispInformationMessage strError
%>
<BR>
<form action="JavaScript:window.history.back()">
	<input type=submit value=" ��  �� ">
</form>
<%
		End If
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
