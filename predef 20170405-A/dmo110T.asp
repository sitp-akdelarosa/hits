<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo110T.asp				_/
'_/	Function	:空搬入情報一覧画面トップ		_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<%
'ユーザデータ所得
        dim USER, COMPcd, LinUN
        USER   = Session.Contents("userid")
        COMPcd = Session.Contents("COMPcd")
	LinUN  = Session.Contents("LinUN")
	'2009/02/25 Add-S G.Ariola	
	Session("Key1") = ""
	Session("Key2") = ""
	Session("Key3") = ""
	
	Session("KeySort1") = ""
	Session("KeySort2") = ""
	Session("KeySort3") = ""
	'2009/02/25 Add-E G.Ariola	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空搬入情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<!--2008/01/29 Add-S M.Marquez-->
<SCRIPT src="./JS/KeyDown.js"></SCRIPT>
<!--2008/01/29 Add-E M.Marquez-->
<SCRIPT language=JavaScript>
<!--
//検索
function Serch(){
  target = document.serch;
  Num=LTrim(target.SortKye.value);
  if(Num.length==0){
    alert("コンテナ番号を記入してください");
    target.SortKye.focus();
    return false;
  }
  if(!CheckEisu(target.SortKye.value)){
    alert("コンテナ番号に半角英数字と半角スペース、「-」、「/」以外の文字を記入しないでください");
    target.SortKye.focus();
    return false;
  }
  parent.DList.SerchC("4",target.SortKye.value);
}
//ソート
function sort(){
  target = document.serch;
  target.SortFlag.value=target.Sort.options[target.Sort.selectedIndex].value;
  target.target="DList";
  target.action="./dmo110L.asp";
  target.submit();
}
//2008-01-29 Add-S M.Marquez
function finit(){
//    document.serch.Sort.focus();
document.serch.SortKye.focus();
}
//2008-01-29 Add-E M.Marquez
function OpenCodeWin()
{
	var CodeWin;
	var w=400;
	var h=300;
	var l=0;
	var t=0;
	if(screen.width){
		l=(screen.width-w)/2;
	}
	if(screen.availWidth){
		l=(screen.availWidth-w)/2;
	}
	if(screen.height){
		t=(screen.height-h)/2;
	}
	if(screen.availHeight){
		t=(screen.availHeight-h)/2;
	}
	
  CodeWin = window.open("./sort.asp?user=<%=Session.Contents("userid")%>&left_menu=2","codelist","scrollbars=yes,resizable=yes,width="+w+",height="+h+",top="+t+",left="+l);
  CodeWin.focus();

}

function showContent(){
    var target1 = document.getElementById("loading");
    target1.style.display='block';
    //show content	
    //parent.DList.document.getElementById("content").style.display='block';
}
// -->
</SCRIPT>
<style>
TD.bordering
{
    BORDER-BOTTOM: 1px dotted #000000;
    BORDER-LEFT: 1px dotted #000000;
    BORDER-RIGHT: 1px dotted #000000;
    BORDER-TOP: 1px dotted #000000;
	
}
</style>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="finit();">
<!-------------空搬入情報一覧画面Top--------------------------->
<TABLE border=1 cellPadding=3 cellSpacing=0 align=center width="100%">
   <TR>
     <TD width="15%">ログインユーザ</TD>
     <TD><%=LinUN%></TD>
     <TD width="7%"><%=USER%></TD>
     <TD width="5%"><%=COMPcd%></TD></TR>
</TABLE>
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <FORM name="serch" action="">
   <TR><TD width="60%"><B class=title>空搬入作業一覧</B><BR><BR>
      
<%'CW-024 Dell IF SortFlag <> 6 Then %>
     <SELECT name="Sort" onChange="sort();">
       <OPTION value=0>搬入予定日順に表示</OPTION>
       <OPTION value=1>指示先が未回答のコンテナ一覧</OPTION>
       <OPTION value=7>指示先回答がNoのコンテナ一覧</OPTION>
       <OPTION value=2>搬入未完了分をすべて表示</OPTION>
       <OPTION value=3>全件表示</OPTION>
     </SELECT>
<%'CW-024 Dell End If %>

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=button value="並べ替え" OnClick="OpenCodeWin()">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=button style="WIDTH: 120px;" value="空搬入データ更新" OnClick="parent.parent.Top.location.reload(true);sort();">
<!--<div class="right" id="loading" >しばらくお待ちください。&nbsp;<IMG border=0 src=Image/loaded.gif></div> -->
<!--2009/07/16 Upd-S G.Ariola -->
<!--<TD align="left"  width="30%" class="bordering"> -->
<TD align="left"  width="30%">
<!--2009/07/16 Upd-E G.Ariola -->
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
<TR><TD>
     <TD>
	<INPUT type=hidden name=SortFlag value="" >
	コンテナ番号:<BR>
	<INPUT type=text  name="SortKye" maxlength=20 size=27>
	<INPUT type=button value="検索" onClick="Serch()">
</TD><TR>	
</TABLE>
	</TD>
<TD align="left"  width="10%">&nbsp;</TD>
	<TR>
  </FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
