<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �w������̎擾�i�s�����ʏ��j
    Dim strLinkUrl
    Dim strLogId
    Dim strLogNo
    Dim strLinkName

    strLogId = Request.QueryString("longid")
    strLogNo = Request.QueryString("logno")
    strLinkName = Request.QueryString("linkname")
    strLinkUrl = Request.QueryString("linkurl")

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    If strLogId = "" or strLogNo = "" or strLinkName = "" Then
        strLogId = "l999"
        strLogNo = "01"
        strLinkName = "���̑�"
    End If

    ' �����N�����o��
    WriteLog fs, strLogId, strLinkName, strLogNo, ","

    ' �s�����ʂփ��_�C���N�g
    Response.Redirect strLinkUrl
%>
