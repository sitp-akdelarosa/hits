<%
	@LANGUAGE = VBScript
	@CODEPAGE = 932
%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi000M.asp				_/
'_/	Function	:事前情報一覧画面メニュー		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:3th   2004/01/31	3次対応		_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
//遷移
function opnewin(i){
  Fname=document.dmi000MF;
  Fname.target="List";
  switch(i){
    case 1:
        Fname.action="./dmo010F.asp";
        break;
    case 2:
        Fname.action="./dmo110F.asp";
        break;
    case 3:
        Fname.action="./dmo210F.asp";
        break;
    case 4:
        Fname.action="./dmo310F.asp";
        break;
    case 5:
      	Win = window.open('dmi010.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
    case 6:
      	Win = window.open('dmi110.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
    case 7:
      	Win = window.open('dmi210.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
    case 8:
      	Win = window.open('dmi310.asp', 'FConIn', 'width=200,height=400,resizable=yes,scrollbars=yes');
        break;
	case 9:
        w=625;
        h=375;
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
    	Win = window.open("dmi410.asp", "FConIn", "width="+w+",height=" + h +",top="+t+",left="+l+",resizable=yes,scrollbars=no");
    	break;
	case 10:
        Fname.action="./top.asp";
        break;
  }
  if(i<5 || i == 10){
    Fname.submit();
  }
}
-->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY class="menu">
<!-------------事前入力初期画面Menu--------------------------->
<CENTER>
<P><B><Font color="#000066">事前情報</FONT></B></P>
<P><A HREF="JavaScript:opnewin(10)">作業<BR>テーブル</A></P>
<P><BR></P>
<P><B><Font color="#000066">各種一覧画面</FONT></B></P>
<P><A HREF="JavaScript:opnewin(1)">実搬出</A></P>
<P><A HREF="JavaScript:opnewin(2)">空搬入</A></P>
<P><A HREF="JavaScript:opnewin(3)">空搬出</A></P>
<P><A HREF="JavaScript:opnewin(4)">実搬入</A></P>
<P><B><Font color="#000066">各種入力画面</FONT></B></P>
<% If Session.Contents("UType") = 3 Then 
     Response.Write "<P>実搬出</P>"
   Else
     Response.Write "<P><A HREF='JavaScript:opnewin(5)'>実搬出</A></P>"
   End If %>
<P><A HREF="JavaScript:opnewin(6)">空搬入</A></P>
<P><A HREF="JavaScript:opnewin(7)">空搬出</A></P>
<P><A HREF="JavaScript:opnewin(8)">実搬入</A></P>
<P><A HREF="JavaScript:opnewin(9)">作業発生<BR>mail設定</A></P>
<FORM name="dmi000MF">
</FORM>
</CENTER>
<!-------------画面終わり--------------------------->
</BODY></HTML>
