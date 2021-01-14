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
<title>Windows 2000 インターネット サービスへようこそ</title>
<style>
	ul{margin-left: 15px;}
	.clsHeading {font-family: ＭＳ Ｐゴシック; color: black; font-size: 11; font-weight: 800; width:210;}	
	.clsEntryText {font-family: ＭＳ Ｐゴシック; color: black; font-size: 11; font-weight: 400; background-color:#FFFFFF;}		
	.clsWarningText {font-family: ＭＳ Ｐゴシック; color: #B80A2D; font-size: 11; font-weight: 600; width:550;  background-color:#EFE7EA;}	
	.clsCopy {font-family: ＭＳ Ｐゴシック; color: black; font-size: 11; font-weight: 400;  background-color:#FFFFFF;}	
	
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
    BORDER="0" ALIGN="LEFT" vspace="0" hspace="0"> <strong>現在、ユーザー向けの既定のドキュメントが設定されていません。
    このサイトに接続しようとしているユーザーには以下のページが表示されます <a
    href="<%= "iisstart.asp?uc=1" %>">"工事中のページ"</a>。</strong>

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
	IIS 5.0 へようこそ</span><br>
    	<span CLASS="clsEntryText">		
	Microsoft Windows 2000 のインターネット インフォメーション サービス (IIS) によって、Windows を強力な Web サーバーとして使用できます。ファイルやプリンタを共有したり、情報を安全に公開するためのアプリケーションを簡単に作成することができます。 IIS は、電子商取引ソリューションを構築、提供するための安全なプラットフォームです。また、Web 上での重要なビジネス アプリケーションの導入が容易になります。
	<P>
	IIS は次のようなニーズに応えることができます。:</span>
	<p>
	<ul class="clsEntryText">
	<li>個人の Web サーバーをセットアップする。
	<li>部署内で情報を共有する。
	<li>データベースにアクセスする。
	<li>企業イントラネットを作成する。
	</ul>
	<p>
	<span CLASS="clsEntryText">
	IIS はインターネット標準と Windows を統合しているので、Web の発行、管理、開発について新たに習得する必要がありません。
	<P>
	Windows 2000 とインターネット インフォメーション サービスは、
	Web 上で情報を共有したり、アプリケーションを実行するのに最も簡単な方法を提供します。
	</span>
	</td>

		<td valign="top">
			<IMG SRC="mmc.gif">
		</td>
		<td valign="top">
			<span CLASS="clsHeading">統合管理</span>
			<br>
			<span CLASS="clsEntryText">
				Windows 2000 の [コンピュータの管理]、<a href="javascript:activate();">コンソール</a>、またはスクリプトを使用して IIS を管理することができます。Windows 2000 Server または Windows 2000 Advanced Server がインストールされている場合は、
			<% if port <> "" then %><A HREF="<%=adminURL%>">管理者 Web サイト</A><% else %>管理者 Web サイト<% end if %> を使用することもできます。 
			<p>
			フォルダを右クリックして、
			一般的な IIS の設定を構成するとともに、Web 経由でコンテンツを共有することもできます。
			</span>
		</td>
	</tr>
	<tr>
		<td valign="top">
			<IMG SRC="help.gif">
		</td>
		<td valign="top">
			<span CLASS="clsHeading"><a href="javascript:loadHelpFront();">オンライン マニュアル</a></span>
			<br>
			<span CLASS="clsEntryText">IIS のオンライン マニュアルには、索引、
 			   検索、およびトピックごとの印刷機能などが含まれます。また、次のようなことが可能です:<p>
			</span>
			<ul class="clsEntryText">
 		 		<li>さまざまなタスクやサーバーの操作についての説明を参照する。
				<li>コード リファレンスを参照する。
		 		<li>コード サンプルを表示する。
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
			<span CLASS="clsHeading">Web 印刷</span>
			<br>
			<span CLASS="clsEntryText">Windows 2000 はサーバー上のプリンタを、アクセスしやすい
 			   <a HREF="/printers" target="_new">Web サイト</a>に動的に一覧表示します。 このサイトで
 			   プリンタとその印刷ジョブを監視できます。また、このサイトを通じてプリンタに他の Windows コンピュータ
			から接続することがができます。</span>
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

