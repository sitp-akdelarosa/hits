<%@ LANGUAGE="VbScript" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<!--#include file="XlsCrt3vbs.inc"-->
<HTML>
<HEAD>
<TITLE>����`�[�t�@�C���o��</TITLE>
</HEAD>
<BODY>
<%
	On Error Resume Next

	wDate = FormatDateTime (Date,vbLongDate)
	wTime = FormatDateTime (Time,vbLongTime)
	wIPAddress    = Request.ServerVariables("REMOTE_ADDR")

	wOutFileName = "Xls3_Uriage" & wDate & Replace(wTime,":","_") & wIPAddress & ".xls"

	'---------------------------------------------------------
	' �ȉ���wFilePath�AwIISFilePath�AwInFileName�̒l�́A
        ' ���s���̍\���ɏ]���ύX���ĉ�����
	'---------------------------------------------------------
	wFilePath  = "c:\inetpub\wwwroot\outfiles\"
	wIISFilePath = "http://localhost/outfiles/"
	wInFileName = "c:\Xls3Sample.xls"

	

	'--------------------------------------------------------
	'�@���͓��e��ϐ��Ɏ擾
	'--------------------------------------------------------
	wUDate    =  Request("UDate")    '�����
	wUNo      =  Request("UNo")      '�`�[No
	wTName    =  Request("TName")    '���Ӑ於
	wTAddress =  Request("TAddress") '���Ӑ�Z���i�[�i��j
	wShimei   =  Request("Shimei")   '�����i�[�i��j
	wSCode    =  Request("SCode")    '���i�R�[�h
	wSName    =  Request("SName")    '���i��
	wSuu      =  Request("Suu")      '����
	wTanka    =  Request("Tanka")    '�P��
	
	'--------------------------------------------------------
	'  ExcelCreator �I�u�W�F�N�g������Excel�t�@�C���o��
	'--------------------------------------------------------
        'ExcelCreator �I�u�W�F�N�g����
        Set Xls1= Server.CreateObject("ExcelCrtOcx.ExcelCrtOcx.1")  

	'����`�[(�I�[�o�[���C)�t�@�C���I�[�v��
  	Xls1.OpenBook wFilePath & wOutFileName,wInFileName

        '���^�V�[�g���Ăяo��
        Xls1.SheetNo = 0

    '�u���E�U��œ��͂����f�[�^���V�[�g�ɏo��
	Xls1.Cell("**UDate").Str    = wUDate         '�����
	Xls1.Cell("**UNo").Str      = wUNo           '�`�[No
	Xls1.Cell("**TName").Str    = wTName & "�l"  '���Ӑ於
	Xls1.Cell("**TAddress").Str = wTAddress      '���Ӑ�Z���i�[�i��j
	Xls1.Cell("**Shimei").Str   = wShimei & "�l" '�����i�[�i��j
	Xls1.Cell("**SCode").Str    = wSCode         '���i�R�[�h
	Xls1.Cell("**SName").Str    = wSName         '���i��
	Xls1.Cell("**Suu").Long     = CLng(wSuu)     '����
	Xls1.Cell("**Tanka").Double = CDbl(wTanka)   '�P��


	wGoukei = CLng(wSuu) * CDbl(wTanka) '���v���z�i�Ŕ����j
	Xls1.Cell("**Kingaku").Double = CDbl(wGoukei)

	wZei = wGoukei * 100 * 0.05 / 100  '�Ŋz
	Xls1.Cell("**Zei").Value = wZei
	
	Xls1.Cell("I18").Func2 "=SUM(C18,F18)",wGoukei + wZei   '�ō��ݍ��v���z��

	wMsg = "Excel�t�@�C�����쐬���܂����B�ȉ����쐬�����t�@�C�����_�E�����[�h�ł��܂�"

	wErrNo = Xls1.ErrorNo
	If wErrNo <> 0 Then
		wMsg = "ExcelCreator3�G���[���b�Z�[�W�F" & Xls1.ErrorMessage
	End If
	Xls1.CloseBook

        Set Xls1 = Nothing
%>
<FONT SIZE="2"><%=wMsg%></FONT><BR>
<% If wErrNo = 0 Then %>
    <Font Size="2">���������t�@�C���̃_�E�����[�h</font><br>
    <Font Size="2"><a href="<%=wIISFilePath%><%=wOutFileName%>"><%=wOutFileName%></A></font>
<% End If %>
</BODY>
</HTML>