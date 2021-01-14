<% @Language = "VBScript" %>
<% Response.buffer = true %>
<!--
	  WARNING!
	  Please do not alter this file. It may be replaced if you upgrade your web server 
     If you want to use it as a template, we recommend renaming it, and modifying the new file.
	  Thanks.
-->


<HTML>

<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text-html; charset=shift_jis">

	
<%
Dim strServername, strLocalname, strServerIP

strServername = LCase(Request.ServerVariables("SERVER_NAME"))
strServerIP = LCase(Request.ServerVariables("LOCAL_ADDR"))
strRemoteIP =  LCase(Request.ServerVariables("REMOTE_ADDR"))


%>
<% If Request("uc") <> 1 AND  (strServername = "localhost" or strServerIP = strRemoteIP) then %>
	<% Response.Redirect "localstart.asp" %>
<% else %>
<title id=titletext>工事中</title>
</HEAD>
	<body bgcolor=white>
	<TABLE>
	<TR>
	<td id="tableProps" width=70 valign=top align=center>
	<IMG id="pagerrorImg" SRC="pagerror.gif" width=36 height=48>  
	<TD id="tablePropsWidth" width=400>
	
	<h1 id=errortype style="font:14pt/16pt ＭＳ Ｐゴシック; color:#4e4e4e">
	<id id="Comment1"><!--Problem--></id><id id="errorText">工事中</id></h1>
	<id id="Comment2"><!--Probable causes:<--></id><id id="errordesc"><font style="font:9pt/12pt ＭＳ Ｐゴシック; color:black">
	接続しようとしたサイトには現在、既定のページがありません。更新中の可能性があります。
	</id>
	<br><br>
	
	<hr size=1 color="blue">
	
	<br>
	<ID  id=term1>
	しばらくしてからこのサイトにもう一度アクセスしてください。問題が続くようであれば Web サイト管理者に連絡してください。
	</ID>
	<P>
	
	</ul>
	<BR>
	</TD>
	</TR>
	</TABLE>
	</BODY>
<% end if %>

</HTML>










