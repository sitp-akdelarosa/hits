<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<!--#include file="MS-ExpCom.inc"-->

<%
    ' Temp�t�@�C�������̃`�F�b�N
    CheckTempFile "MSEXPORT", "expentry.asp"

    ' ���[�U��ނ��`�F�b�N����
    strUserKind=Session.Contents("userkind")

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
    CsvTitleWrite strUserKind

    '�A�o�R���e�iCSV�t�@�C���f�[�^�s�o��
    CsvDataWrite strUserKind, ti
 
   ' �A�o�R���e�i�Ɖ�
    WriteLog fs, "1109","�A�o�R���e�i�Ɖ�-CSV�t�@�C���o��","30", ","

    ' �_�E�����[�h�I��
    Response.End

%>
