<%@Language="VBScript" %>

<!--#include file="./Common/Common.inc"-->

<%
	'�ϐ��錾
	Dim strFileName

	' Temp�t�@�C�������̃`�F�b�N

	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' �_�E�����[�h�t�@�C���̎擾

	strFileName = Session.Contents("tempfile")
	If strFileName="" Then
		' �Z�b�V�������؂�Ă���Ƃ�
		Response.Redirect "accesstotal.asp"	 '���p����Top��
		Response.End
	End If
	strFileName="../temp/" & strFileName
	' �_�E�����[�h�t�@�C����Open
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' �t�@�C���̃_�E�����[�h
	Response.ContentType="application/octet-stream"
	Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

	'�w�b�_��������
	Response.Write "�A�N�Z�X�����݌v�\"
	Response.Write Chr(13) & Chr(10)
	Response.Write Chr(13) & Chr(10)

	'CSV�t�@�C���^�C�g���s�o��
	Response.Write "�敪,"
	Response.Write "PC,"
	Response.Write "�g�ђ[��,"
	'Y.TAKAKUWA Add-S 2013-09-30
	Response.Write "�`�Ԓ[��,"
	'Y.TAKAKUWA Add-E 2013-09-30
	Response.Write "���v," 
	Response.Write "�݌v" 
	Response.Write Chr(13) & Chr(10)
	
	'�݌vCSV�t�@�C���f�[�^�s�o��
	Do While Not ti.AtEndOfStream
		anyTmp=Split(ti.ReadLine,",")
		Response.Write anyTmp(0) & ","
		Response.Write anyTmp(1) & ","
		Response.Write anyTmp(2) & ","
		'2013-09-30 Y.TAKAKUWA Upd-S
		'Response.Write anyTmp(3) & ","
		'Response.Write anyTmp(4) & ""
		Response.Write anyTmp(3) & ","
		Response.Write anyTmp(4) & ","
		Response.Write anyTmp(5) & ""
		'2013-09-30 Y.TAKAKUWA Upd-E
		Response.Write Chr(13) & Chr(10)
	Loop


	' �_�E�����[�h�I��
	Response.End

%>
