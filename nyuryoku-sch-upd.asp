<%@Language="VBScript" %>

<!--#include file="Common.inc"-->
<!--#include file="vessel.inc"-->

<%
    ' �Z�b�V�����̃`�F�b�N
    CheckLogin "nyuryoku-in1.asp"

    ' �G���[�t���O�̃N���A
    bError = false

    ' ���̓t���O�̃N���A
    bInput = true

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �\���t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �����w��̂Ȃ��Ƃ�
        strFileName="test.csv"
    End If
    strFileName="./temp/" & strFileName
    ' �\���t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

'�e�����ݒ�
    Dim sAdate, sTdate, sDdate, sCdate, sRdate
	'���ݗ\�莞��
	sAdate = ""
	If Request.Form("ayear") <> "" Then
	    saDate = SetDateTime(Request.Form("ayear"), Request.Form("amonth"), Request.Form("aday"), _ 
	                         GetNumStr(Request.Form("ahour"), 2), GetNumStr(Request.Form("amin"), 2))
	End If
	'���݊�������
	sTdate = ""
	If Request.Form("tyear") <> "" Then
	    sTdate = SetDateTime(Request.Form("tyear"), Request.Form("tmonth"), Request.Form("tday"), _ 
	                         GetNumStr(Request.Form("thour"), 2), GetNumStr(Request.Form("tmin"), 2))
	End If
	'���݊�������
	sDdate = ""
	If Request.Form("dyear") <> "" Then
	    sDdate = SetDateTime(Request.Form("dyear"), Request.Form("dmonth"), Request.Form("dday"), _ 
	                         GetNumStr(Request.Form("dhour"), 2), GetNumStr(Request.Form("dmin"), 2))
	End If
	'����Long Schedule
	sCdate = ""
	If Request.Form("cyear") <> "" Then
	    sCdate = SetDateTime(Request.Form("cyear"), Request.Form("cmonth"), Request.Form("cday"), _ 
	                         "23", "59")
'	                         GetNumStr(Request.Form("chour"), 2), GetNumStr(Request.Form("cmin"), 2))
	End If
	'����Long Schedule
	sRdate = ""
	If Request.Form("ryear") <> "" Then
	    sRdate = SetDateTime(Request.Form("ryear"), Request.Form("rmonth"), Request.Form("rday"), _ 
	                         "23", "59")
'	                         GetNumStr(Request.Form("rhour"), 2), GetNumStr(Request.Form("rmin"), 2))
	End If

    ' �w������̎擾(�w��s)
    Dim iLine, sIn1, sIn2, sInpFlg
	Dim sText(35) 

    iLine = Cint(Trim(Request.QueryString("line")))


    ' �ڍו\���s�̃f�[�^�̎擾

    Dim iKensu		'�\������(��ʕ\������)
    Dim LineNo		'�t�@�C���̃��C���J�E���^
    Dim iHitNo		'��v����t�@�C���s��
	Dim iDelLine	'�폜����s�ԍ�

    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
		Select Case LineNo
			Case 1
				iKensu = anyTmp(7)
		    	iHitNo = 2 +  iLine
				sText(LineNo) = anyTmp(0) &  "," & _
							    anyTmp(1) &  "," & _
							    anyTmp(2) &  "," & _
							    anyTmp(3) &  "," & _
							    anyTmp(4) &  "," & _
							    anyTmp(5) &  "," & _
							    anyTmp(6) &  "," & iKensu
			Case 2
				sText(LineNo) = iKensu
			Case Else
		        If LineNo = iHitNo Then
					sText(LineNo) = anyTmp(0) &  "," & anyTmp(1) &  "," & _
								    saDate    &  "," & sTdate    &  "," & _
							    	anyTmp(4) &  "," & sDdate    &  "," & _
							    	sCdate &  "," & sRdate
				Else
					sText(LineNo) = anyTmp(0) &  "," & anyTmp(1) &  "," & _
								    anyTmp(2) &  "," & anyTmp(3) &  "," & _
								    anyTmp(4) &  "," & anyTmp(5) &  "," & _
							    	anyTmp(6) &  "," & anyTmp(7) 
		        End If
		End Select
    Loop
    ti.Close

'���ԕ��ёւ��̏������s��(��������̗v�]�ŁA�R�����g�� 2002/02/27)
'*** Start M.Hayashi ****
'	Dim sBefDate
'	Dim sAftDate
'   Dim sWkText
'   Dim bSwap
'   For i = 3 to LineNo - 1
'		anyTmp=Split(sText(i),",")
'		sBefDate = anyTmp(2)
'		For j = (i + 1) To LineNo
'           anyTmp=Split(sText(j),",")
'		    sAftDate = anyTmp(2)
'           bSwap = FALSE
'           If sAftDate <> "" Then
'			  If sBefDate = "" Then
'               bSwap = TRUE
'             Else
'               If sBefDate > sAftDate Then
'                 bSwap = TRUE
'				End If
'             End IF
'           End If
'           If bSwap = TRUE Then
'             sWkText = sText(i)
'             sText(i) = sText(j)
'             sText(j) = sWkText
'           End IF
'		Next 
'	Next
'*** End   M.Hayashi ****

    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),2,True)
    For i = 1 to LineNo
		ti.WriteLine sText(i)
    Next
    ti.Close

	sCdate = Left(sCdate,10)
	sRdate = Left(sRdate,10)
    ' �{�D���ÍX�V����
    WriteLog fs, "3004","�D�Ё^�^�[�~�i������-�{�D���Ó���","12", Request.Form("sportname") & "/" & saDate & "/" & sTdate & "/" & sDdate & "/" & sCdate & "/" & sRdate & ","

%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body>
<!-------------��������ꗗ���--------------------------->
<!-------------�o�^��ʍX�V�����I���--------------------------->
</body>

<SCRIPT LANGUAGE="JavaScript">
	window.location.replace("nyuryoku-port.asp");
</SCRIPT>
</html>
