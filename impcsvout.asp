<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="ImpCom.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "IMPORT", "impentry.asp"

    ' �\�����[�h�̎擾
    Dim bDispMode          ' true=�R���e�i���� / false=BL����
    If Session.Contents("findkind")="Cntnr" Then
        bDispMode = true
        strOption = "�R���e�iNo.CSV�t�@�C�����M"
    Else
        bDispMode = false
        strOption = "BL�ԍ�CSV�t�@�C�����M"
    End If

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' �_�E�����[�h�t�@�C���̎擾
    Dim strFileName
    strFileName = Session.Contents("tempfile")
    If strFileName="" Then
        ' �Z�b�V�������؂�Ă���Ƃ�
        Response.Redirect "impentry.asp"             '�A���R���e�i�Ɖ�g�b�v
        Response.End
    End If
    strFileName="./temp/" & strFileName
    ' �_�E�����[�h�t�@�C����Open
    Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

    ' �t�@�C���̃_�E�����[�h
    Response.ContentType="application/octet-stream"
    Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

    '�A���R���e�iCSV�t�@�C���^�C�g���s�o��
    CsvTitleWrite bDispMode

    '�A���R���e�iCSV�t�@�C���f�[�^�s�o��
    CsvDataWrite bDispMode, ti

    ' �A���R���e�i�Ɖ�
    WriteLog fs, "2008","�A���R���e�i�Ɖ�-CSV�t�@�C���o��","30", filename & ","

    ' �_�E�����[�h�I��
    Response.End

%>
