<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' �w������̎擾�i�s�����ʏ��j
    Dim strLinkID
    strLinkID = Request.QueryString("link")

    ' File System Object �̐���
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    Select Case strLinkID
        Case "http://www.fk-tosikou.or.jp"  strLinkNamne = "�����k��B�������H����"
        Case "http://www.jartic.or.jp"      strLinkNamne = "�i���j���{���H��ʏ��Z���^�["
        Case Else                           strLinkNamne = "�s��"
    End Select

    ' �����N�����o��
	strLogInfo = "�Q�[�g�O�f���E���G�󋵏Љ�-�����N-" & strLinkNamne
    WriteLog fs, "8001",strLogInfo,"01", ","

    ' �s�����ʂփ��_�C���N�g
    Response.Redirect strLinkID
%>
