<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="MS-ImpCom.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSIMPORT", "impentry.asp"

    ' ���[�U��ނ��`�F�b�N����
    strUserKind=Session.Contents("userkind")

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
    CsvTitleWrite strUserKind

    '�A���R���e�iCSV�t�@�C���f�[�^�s�o��
    CsvDataWrite strUserKind, ti

    ' �A���R���e�i�Ɖ�
    WriteLog fs, "2110","�A���R���e�i�Ɖ�-CSV�t�@�C���o��","30", filename & ","

    ' �_�E�����[�h�I��
    Response.End

%>
