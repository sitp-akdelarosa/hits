<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "EXPORT", "bookentry.asp"

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �_�E�����[�h�t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "bookentry.asp"             '�A�o�R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' �_�E�����[�h�t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �t�@�C���̃_�E�����[�h
    Response.ContentType="application/octet-stream"
    Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

    '�A�o�R���e�iCSV�t�@�C���^�C�g���s�o��
    Response.Write "Booking No.,"
    Response.Write "�D��,"
    Response.Write "�D��,"
    Response.Write "Voyage No.,"
    Response.Write "�d���`,"
    Response.Write "��R�����o�ꏊ,"
    Response.Write "CY�J�b�g,"	'I20080222
    Response.Write "�T�C�Y,"
    Response.Write "�^�C�v,"
    Response.Write "����,"
    Response.Write "�ގ�,"	'I20040223
    Response.Write "�\��{��,"
    Response.Write "���o�ϖ{��"

    Response.Write Chr(13) & Chr(10)

    '�A�o�R���e�iCSV�t�@�C���f�[�^�s�o��
    Do While Not ti.AtEndOfStream
        anyTmp=Split(ti.ReadLine,",")

        Response.Write anyTmp(1) & ","
        Response.Write anyTmp(2) & ","
        Response.Write anyTmp(3) & ","
        Response.Write anyTmp(4) & ","
        Response.Write anyTmp(5) & ","
        Response.Write anyTmp(6) & ","
        Response.Write anyTmp(14) & ","		'I20080222
        Response.Write anyTmp(7) & ","
        Response.Write anyTmp(8) & ","
        Response.Write anyTmp(9) & ","
        Response.Write anyTmp(12) & ","		'I20040223
        Response.Write anyTmp(10) & ","
        Response.Write anyTmp(11)

'		If UBound(anyTmp)>11 Then
''			For i=12 To UBound(anyTmp)	'D20040223
'			For i=13 To UBound(anyTmp)	'I20040223
'				Response.Write "," & anyTmp(i)
'			Next
'		End If

        Response.Write Chr(13) & Chr(10)
    Loop

   ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "1013","�u�b�L���O���Ɖ�-CSV�t�@�C���o��","30", ","

    ' �_�E�����[�h�I��
    Response.End

%>
