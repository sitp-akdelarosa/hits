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

    ' �w������̎擾(�w��s)
    Dim iLine, sIn1, sIn2, sInpFlg
	Dim sText(35) 

    iLine = Trim(Request.QueryString("line"))

    ' �ڍו\���s�̃f�[�^�̎擾

    Dim iKensu		'�\������(�폜��)
    Dim LineNo		'�t�@�C���̃��C���J�E���^
    Dim iHitNo		'��v����t�@�C���s��
	Dim iDelLine	'�폜����s�ԍ�
	Dim sPortName	'���O�p�`��

    LineNo=0
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")
        LineNo=LineNo+1
		Select Case LineNo
			Case 1
				iKensu = anyTmp(7) - 1
				If Cint(iKensu) <> 0 Then
		    		iHitNo = 2 +  Cint(iLine)
				End If
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
				If iKensu = 0 Then
					Exit Do
				End If

		        If LineNo = iHitNo Then
					iDelLine = LineNo
					sPortName = anyTmp(1)
				Else
					sText(LineNo) = anyTmp(0) &  "," & _
								    anyTmp(1) &  "," & _
								    anyTmp(2) &  "," & _
								    anyTmp(3) &  "," & _
								    anyTmp(4) &  "," & _
							    	anyTmp(5) &  "," & _
							    	anyTmp(6) &  "," & _
							    	anyTmp(7)
		        End If
		End Select
    Loop
    ti.Close


	sBk = Server.MapPath(strFileName)
	sTemp  = strFileName & ".tmp" 	'Server.MapPath(strFileName)
	fs.deletefile sBk, True				'��x�폜
	ti  = Server.MapPath(sTemp)
    Set ti=fs.OpenTextFile(Server.MapPath(sTemp),2,True)
    For i = 1 to LineNo

		If iKensu <> 0 Then
			If iDelLine <> i Then
				ti.WriteLine sText(i)
			End If
		Else
			ti.WriteLine sText(i)
			If i = 2 Then
				Exit For
			End If
		End If
    Next
	ti  = Server.MapPath(sTemp)
	sBk = Server.MapPath(strFileName)
	fs.MoveFile ti, sBk



    ' �{�D���Í폜
    WriteLog fs, "3004","�D�Ё^�^�[�~�i������-�{�D���Ó���", "13", sPortName & ","

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
