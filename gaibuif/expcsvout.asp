<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ExpCom.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "EXPORT", "expentry.asp"

    ' �\�����[�h�̎擾
    Dim bDispMode          ' true=�R���e�i���� / false=�u�b�L���O����
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
        strOption = "�R���e�iNo.CSV�t�@�C�����M"
    Else
        bDispMode = false
        strOption = "Booking�ԍ�CSV�t�@�C�����M"
    End If

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �_�E�����[�h�t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "expentry.asp"             '�A�o�R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' �_�E�����[�h�t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �t�@�C���̃_�E�����[�h
    Response.ContentType="application/octet-stream"
    Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

    '�A�o�R���e�iCSV�t�@�C���^�C�g���s�o��
    CsvTitleWrite bDispMode

    '�A�o�R���e�iCSV�t�@�C���f�[�^�s�o��
    CsvDataWrite bDispMode, ti

    ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "3008","�d�o�n���Ɖ�-CSV�t�@�C���o��","30", ","

    ' �_�E�����[�h�I��
    Response.End

%>
