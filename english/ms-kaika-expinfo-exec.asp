<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="Vessel.inc"-->

<%
	'
	'	�y�C�ݓ��́z	�G���[�`�F�b�N�A�\���A�t�@�C���쐬
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
        Response.Redirect "ms-kaika-expinfo-updatecheck.asp"

	ElseIf bKind=2 And sStop<>"" Then
        Response.Redirect "ms-kaika-expinfo-list.asp"

	Else
		' �g�����U�N�V�����t�@�C���̊g���q 
		Const SEND_EXTENT = "snd"
		' �g�����U�N�V�����h�c
		Const sTranID = "EX16"
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
	    Dim sUser,sUserNo,sVslCode,sVoyCtrl,sBooking,sTraderCode,iSize,sType,iHeight,sPick,sEmpDate,sCyDate
		Dim sEmpDateS,sCyDateS,sShipLine
	    sUser 	= UCase(Trim(Request.form("user")))
	    sUserNo = UCase(Trim(Request.form("userno")))
	    sVslCode = UCase(Trim(Request.form("vslcode")))
	    sVoyCtrl = UCase(Trim(Request.form("voyctrl")))
	    sBooking = UCase(Trim(Request.form("booking")))
	    sTraderCode = UCase(Trim(Request.form("tradercode")))
	    iSize 	= Trim(Request.form("size"))
	    sType 	= UCase(Trim(Request.form("type")))
	    iHeight = Trim(Request.form("height"))
	    sPick 	= UCase(Trim(Request.form("pickplace")))
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
		sCyDate = Trim(Request.form("cydate_year")) 
		sCyDate = sCyDate & Right("0" & Trim(Request.form("cydate_mon")),2)
		sCyDate = sCyDate & Right("0" & Trim(Request.form("cydate_day")),2)
		If sCyDate=00 Then
			sCyDate = ""
		Else
			sCyDateT = Trim(Request.form("cydate_year")) 
			sCyDateT = sCyDateT & "/" & Right("0" & Trim(Request.form("cydate_mon")),2)
			sCyDateT = sCyDateT & "/" & Right("0" & Trim(Request.form("cydate_day")),2)
			sCyDateS = Trim(Request.form("cydate_year")) 
			sCyDateS = sCyDateS & "�N " & Trim(Request.form("cydate_mon"))
			sCyDateS = sCyDateS & "�� " & Trim(Request.form("cydate_day")) & "�� "
		End If

	    ' File System Object �̐���
	    Set fs=Server.CreateObject("Scripting.FileSystemobject")

		' ���p�J���}�`�F�b�N
		If InStr(sVslCode,",")<>0 Or _
			InStr(sVoyCtrl,",")<>0 Or _
			InStr(sBooking,",")<>0 Or _
			InStr(sTraderCode,",")<>0 Or _
			InStr(sRemark,",")<>0 Or _
			InStr(sPick,",")<>0 Or _
			InStr(sUser,",")<>0 Or _
			InStr(sUserNo,",")<>0 _
		Then

		    bError = true
			strError = "���͂̍ہA���p�J���}�͎g�p���Ȃ��ŉ������B"

		Else

			ConnectSvr conn, rsd
			' �׎�R�[�h�Ɖ׎�Ǘ��ԍ��̏d���`�F�b�N
			Dim iRecCount
			If Not bKind=0 Then
				sql = "SELECT count(*) FROM ExportCargoInfo WHERE Shipper='" & sUser & "' AND ShipCtrl='" & sUserNo & "'"
				rsd.Open sql, conn, 0, 1, 1
				If Not rsd.EOF Then
					iRecCount = rsd(0)
					If Not iRecCount=0 Then
					    bError = true
						strError = "�׎�R�[�h�Ɖ׎�Ǘ��ԍ����d�����Ă��܂��B"
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
	    Dim sEX16, iSeqNo_EX16, strFileName, sTran, sTusin, sDate
		'�V�[�P���X�ԍ�
		iSeqNo_EX16 = GetDailyTransNo
		'�ʐM�����擾
		sTusin  = SetTusinDate

		sEX16 = iSeqNo_EX16 & "," & sTranID & "," & sSyori & ","  & sTusin & ",Web - " & _
				sSosin & "," & sPlace & "," & sVslCode & "," &  sVoyCtrl & "," & _
				sBooking & "," & sUser & "," & sUserNo & "," & sSosin & "," & _
				sContainer & "," & iSize & "," & sType & "," & iHeight & "," & sRemark & "," & sTraderCode & "," & _
				sEmpDate & "," & sCyDate & "," & sPick
		sFileName = ArrangeNumV(Month(Now), 2) & ArrangeNumV(Day(Now), 2) & iSeqNo_EX16
		strFileName_01 = "./send/" & sFileName & "." & SEND_EXTENT
	    Set ti=fs.OpenTextFile(Server.MapPath(strFileName_01),2,True)
		ti.WriteLine sEX16
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
	            anyTmp(3) = sUserNo
	            anyTmp(4) = sBooking
	            anyTmp(5) = sTraderCode
	            anyTmp(6) = sEmpDateT
	            anyTmp(7) = sCyDateT
	            anyTmp(8) = iSize
	            anyTmp(9) = sType
	            anyTmp(10) = iHeight
	            anyTmp(11) = sRemark
	            anyTmp(12) = sPick


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

				strTemp = sVslCode & "," &  sVoyCtrl & "," & sUser & "," & sUserNo & "," & _
						 sBooking & "," & sTraderCode & "," & sEmpDateT & "," & sCyDateT & "," & _
						 iSize & "," & sType & "," & iHeight & "," & sRemark & "," & sPick

                ti.WriteLine strTemp
	            ti.Close

			End If

		End If

' Temp�����܂�

	' ���O�t�@�C�������o��
	Dim sLogDate,sLogTime
	sLogDate = Trim(Request.form("cydate_year")) & "/"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("cydate_mon")),2) & "/"
	sLogDate = sLogDate & Right("0" & Trim(Request.form("cydate_day")),2)
	sLogTime = Trim(Request.form("emparvtime_year")) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_mon")),2) & "/"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_day")),2) & " "
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_hour")),2) & ":"
	sLogTime = sLogTime & Right("0" & Trim(Request.form("emparvtime_min")),2)
	If sEmpDateT="" Then
		sLogTime = ""
	End If
	If sCyDateT="" Then
		sLogDate = ""
	End If

	strOption = sVslCode & _
				"/" & sVoyCtrl & _
				"/" & sUser & _
				"/" & sUserNo & _
				"/" & sBooking & _
				"/" & sTraderCode & _
				"/" & sLogTime & _
				"/" & sLogDate & _
				"/" & iSize & _
				"/" & sType & _
				"/" & iHeight & _
				"/" & sPick & _
				"/" & sRemark & ","
    If bError Then
		strOption = strOption &	"���͓��e�̐���:1(���)"
    Else
		strOption = strOption & "���͓��e�̐���:0(������)"
    End If

	If bKind=1 Then
		'�V�K
   		WriteLog fs, "4102","�C�ݓ��͗A�o�ݕ����-������", "11", strOption
	ElseIf sDel<>"" Then
   		WriteLog fs, "4102","�C�ݓ��͗A�o�ݕ����-������", "13", strOption
	Else
   		WriteLog fs, "4102","�C�ݓ��͗A�o�ݕ����-������", "12", strOption
	End If

    If Not bError And bKind=0 Then
		Response.Redirect "ms-kaika-expinfo-list.asp"
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
          <td rowspan=2><img src="gif/kaika4t.gif" width="506" height="73"></td>
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
                    <font color="#FFFFFF"><b>�D��</b></font>
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
                  <td bgcolor="#000099" nowrap align=center valign=middle>
					<font color="#FFFFFF"><b>�׎�R�[�h</b></font>
				  </td>
                  <td nowrap>
					<%=sUser%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>�׎�Ǘ��ԍ�</b></font>
                  </td>
                  <td nowrap>
					<%=sUserNo%>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>Booking No.</b></font>
                  </td>
                  <td nowrap>
					<%=sBooking%>
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
                    <font color="#FFFFFF"><b>��R���q�ɓ����w�����</b></font>
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
                    <font color="#FFFFFF"><b>�b�x�����w���</b></font>
                  </td>
                  <td nowrap>
					<% If sCyDate<>"" Then %>
						<%=sCyDateS%>
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
                    <font color="#FFFFFF"><b>����</b></font>
                  </td>
                  <td nowrap>
					<% If iHeight<>"" Then %>
						<%=iHeight%>
					<% Else %>
						<BR>
					<% End If %>
                  </td>
                </tr>

                <tr> 
                  <td bgcolor="#000099" nowrap align=center valign=middle>
                    <font color="#FFFFFF"><b>��R���s�b�N�ꏊ</b></font>
                  </td>
                  <td nowrap>
					<% If sPick<>"" Then %>
						<%=sPick%>
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
<form>
	<input type=button value=" ��  �� " onClick="JavaScript:window.history.back()">
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