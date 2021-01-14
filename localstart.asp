<% @Language = "VBScript" %>
<% Response.buffer = true %>

<html>

<%
Dim oFS,oFSPath
Dim servername,serverinst, path
Dim oDefSite,sDefDoc,sSitePath,sDocName
Dim aDefDoc
Dim success
Dim infoobj, administ
Dim bind,binditems,port,adminURL

adminURL = ""
success = false

Set infoobj=GetObject("IIS://localhost/w3svc/info")
Set administ= GetObject("IIS://localhost/w3svc/" & infoobj.AdminServer)	
bind = administ.ServerBindings(0)(0)
binditems = split(bind,":")
port= binditems(1)
adminURL = "http://localhost:" & port & "/"


Set oFS=CreateObject("Scripting.FileSystemObject")

servername=Request.ServerVariables("SERVER_NAME")
serverinst=Request.ServerVariables("INSTANCE_ID")

path = "IIS://" & servername & "/W3SVC/" & serverinst
Set oDefSite = GetObject(path)

thisURL = oDefSite.ADsPath & "/Root" & Request.ServerVariables("URL")
if instr(thisURL,"localstart.asp") > 0 then
	thisURL =  Mid(thisURL,1,instr(thisURL,"localstart.asp")-2)
end if
Set oDefSiteRoot = GetObject(thisURL)
'Get the default document for this site...
sDefDoc = oDefSite.DefaultDoc
sSitePath = oDefSiteRoot.Path

'parse through the default document string
aDefDocs = split(sDefDoc,",")

'and make sure at least one of them is valid
for each sDocName in aDefDocs
	if oFS.FileExists(sSitePath & "\" & sDocName) then
		if InStr(sDocName,"iisstart") = 0 then
			success = True
			exit for
		end if
	end if
next
%>





<head>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=shift_jis">
<title>Windows 2000 �C���^�[�l�b�g �T�[�r�X�ւ悤����</title>
<style>
	ul{margin-left: 15px;}
	.clsHeading {font-family: �l�r �o�S�V�b�N; color: black; font-size: 11; font-weight: 800; width:210;}	
	.clsEntryText {font-family: �l�r �o�S�V�b�N; color: black; font-size: 11; font-weight: 400; background-color:#FFFFFF;}		
	.clsWarningText {font-family: �l�r �o�S�V�b�N; color: #B80A2D; font-size: 11; font-weight: 600; width:550;  background-color:#EFE7EA;}	
	.clsCopy {font-family: �l�r �o�S�V�b�N; color: black; font-size: 11; font-weight: 400;  background-color:#FFFFFF;}	
	
</style>
</head>

<body TOPMARGIN="3" LEFTMARGIN="3" MARGINHEIGHT="0" MARGINWIDTH="0" BGCOLOR="#FFFFFF"
LINK="#000066" VLINK="#000000" ALINK="#0000FF" TEXT="#000000">
<!-- BEGIN MAIN DOCUMENT BODY --->

<img src="win2000.gif" vspace="0" hspace="0"> 
<table WIDTH="500" CELLPADDING="5" CELLSPACING="3" BORDER="0">
<% if not success and err = 0 then %>
  <tr>
    <td CLASS="clsWarningText" colspan="2">

	<img SRC="warning.gif" WIDTH="40" HEIGHT="40"
    BORDER="0" ALIGN="LEFT" vspace="0" hspace="0"> <strong>���݁A���[�U�[�����̊���̃h�L�������g���ݒ肳��Ă��܂���B
    ���̃T�C�g�ɐڑ����悤�Ƃ��Ă��郆�[�U�[�ɂ͈ȉ��̃y�[�W���\������܂� <a
    href="<%= "iisstart.asp?uc=1" %>">"�H�����̃y�[�W"</a>�B</strong>

	</td>
  </tr>
<% end if %>
  <tr>
	<td>
	<table CELLPADDING="3" CELLSPACING="3" border=0 >
	<tr>
		<td valign="top" rowspan=3>
			<IMG SRC="web.gif">
		</td>	
		<td valign="top" rowspan=3>
	<span CLASS="clsHeading">
	IIS 5.0 �ւ悤����</span><br>
    	<span CLASS="clsEntryText">		
	Microsoft Windows 2000 �̃C���^�[�l�b�g �C���t�H���[�V���� �T�[�r�X (IIS) �ɂ���āAWindows �����͂� Web �T�[�o�[�Ƃ��Ďg�p�ł��܂��B�t�@�C����v�����^�����L������A�������S�Ɍ��J���邽�߂̃A�v���P�[�V�������ȒP�ɍ쐬���邱�Ƃ��ł��܂��B IIS �́A�d�q������\�����[�V�������\�z�A�񋟂��邽�߂̈��S�ȃv���b�g�t�H�[���ł��B�܂��AWeb ��ł̏d�v�ȃr�W�l�X �A�v���P�[�V�����̓������e�ՂɂȂ�܂��B
	<P>
	IIS �͎��̂悤�ȃj�[�Y�ɉ����邱�Ƃ��ł��܂��B:</span>
	<p>
	<ul class="clsEntryText">
	<li>�l�� Web �T�[�o�[���Z�b�g�A�b�v����B
	<li>�������ŏ������L����B
	<li>�f�[�^�x�[�X�ɃA�N�Z�X����B
	<li>��ƃC���g���l�b�g���쐬����B
	</ul>
	<p>
	<span CLASS="clsEntryText">
	IIS �̓C���^�[�l�b�g�W���� Windows �𓝍����Ă���̂ŁAWeb �̔��s�A�Ǘ��A�J���ɂ��ĐV���ɏK������K�v������܂���B
	<P>
	Windows 2000 �ƃC���^�[�l�b�g �C���t�H���[�V���� �T�[�r�X�́A
	Web ��ŏ������L������A�A�v���P�[�V���������s����̂ɍł��ȒP�ȕ��@��񋟂��܂��B
	</span>
	</td>

		<td valign="top">
			<IMG SRC="mmc.gif">
		</td>
		<td valign="top">
			<span CLASS="clsHeading">�����Ǘ�</span>
			<br>
			<span CLASS="clsEntryText">
				Windows 2000 �� [�R���s���[�^�̊Ǘ�]�A<a href="javascript:activate();">�R���\�[��</a>�A�܂��̓X�N���v�g���g�p���� IIS ���Ǘ����邱�Ƃ��ł��܂��BWindows 2000 Server �܂��� Windows 2000 Advanced Server ���C���X�g�[������Ă���ꍇ�́A
			<% if port <> "" then %><A HREF="<%=adminURL%>">�Ǘ��� Web �T�C�g</A><% else %>�Ǘ��� Web �T�C�g<% end if %> ���g�p���邱�Ƃ��ł��܂��B 
			<p>
			�t�H���_���E�N���b�N���āA
			��ʓI�� IIS �̐ݒ���\������ƂƂ��ɁAWeb �o�R�ŃR���e���c�����L���邱�Ƃ��ł��܂��B
			</span>
		</td>
	</tr>
	<tr>
		<td valign="top">
			<IMG SRC="help.gif">
		</td>
		<td valign="top">
			<span CLASS="clsHeading"><a href="javascript:loadHelpFront();">�I�����C�� �}�j���A��</a></span>
			<br>
			<span CLASS="clsEntryText">IIS �̃I�����C�� �}�j���A���ɂ́A�����A
 			   �����A����уg�s�b�N���Ƃ̈���@�\�Ȃǂ��܂܂�܂��B�܂��A���̂悤�Ȃ��Ƃ��\�ł�:<p>
			</span>
			<ul class="clsEntryText">
 		 		<li>���܂��܂ȃ^�X�N��T�[�o�[�̑���ɂ��Ă̐������Q�Ƃ���B
				<li>�R�[�h ���t�@�����X���Q�Ƃ���B
		 		<li>�R�[�h �T���v����\������B
			</ul>

		</td>
	</tr>
<%

		Dim WshShell, ver
		Set WshShell = Server.CreateObject("Wscript.Shell")
		On Error Resume Next
		ver = 0
		ver = WshShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Printers\DisableWebPrinting")
		
%>

<% If ver <> 1 or err <> 0 Then %>
	<tr>
		<td valign="top">
			<IMG SRC="print.gif">
		</td>
		<td valign="top">
			<span CLASS="clsHeading">Web ���</span>
			<br>
			<span CLASS="clsEntryText">Windows 2000 �̓T�[�o�[��̃v�����^���A�A�N�Z�X���₷��
 			   <a HREF="/printers" target="_new">Web �T�C�g</a>�ɓ��I�Ɉꗗ�\�����܂��B ���̃T�C�g��
 			   �v�����^�Ƃ��̈���W���u���Ď��ł��܂��B�܂��A���̃T�C�g��ʂ��ăv�����^�ɑ��� Windows �R���s���[�^
			����ڑ����邱�Ƃ����ł��܂��B</span>
		</td>
	</tr>
<% end if %>
<% err.clear %> 
	</table>
</td>
</tr>
</table>

<P align=center><EM><A href="/iishelp/common/colegal.htm">(C) 
1997-1999 Microsoft Corporation. All rights 
reserved.</A></EM></P></FONT></BODY>

<script LANGUAGE="javascript">
	var gWinheight
	var gDialogsize
	var ghelpwin;
	//launch help
	window.moveTo(5,5);
	gWinheight= 480;
	gDialogsize= "width=640,height=480,left=300,top=50,"
	if (window.screen.height > 600)
	{
<% if not success and err = 0 then %>
		gWinheight= 700;
<% else %>
		gWinheight= 700;
<% end if %>
		gDialogsize= "width=640,height=480,left=500,top=50"
	}
	
	window.resizeTo(600,gWinheight)
	loadHelpFront();

function loadHelpFront(){
	ghelpwin = window.open("http://localhost/iishelp/","Help","status=yes,toolbar=yes,menubar=yes,location=yes,resizable=yes,"+gDialogsize,true);	
}

function activate(){
	window.open("http://localhost/iishelp/iis/htm/core/iisnapin.htm", "SnapIn", 'toolbar=no, left=200, top=200, scrollbars=no, resizeable=no,  width=350, height=350');
}
</script>

</html>

