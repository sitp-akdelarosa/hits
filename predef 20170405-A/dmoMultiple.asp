<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo320.asp				_/
'_/	Function	:事前実搬入入力画面(表示)		_/
'_/	Date		:2003/05/29				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/08/07	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
	Response.AddHeader "Pragma", "no-cache" 
%>
<!--#include File="Common.inc"-->
<!--#include File="../download/download.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->


<%
'セッションの有効性をチェック
  'CheckLoginH
  'WriteLogH "b402", "実搬入事前情報入力","11",""

    dim file1,gerrmsg
  '2010/02/18 M.Marquez Add-A
  if Request.Form("Gamen_Mode")="R" then 
     wReportName="搬入票" 
     wReportID="dmo320" 
     wOutFileName=gfReceiveReportMultiple()
     file1	= server.mappath(gOutFileForder & wOutFileName)
	 if not gfdownloadFile(file1, wOutFileName) then
			wMsg = Replace(gerrmsg,"<br>","\n")
	 end if

  end if
  '2010/02/18 M.Marquez Add-E
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>実搬入情報入力(表示)</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  window.resizeTo(850,690);
  window.focus();
  len = target.elements.length;
  //for (i=0; i<46; i++) target.elements[i].readOnly = true;
  //bgset(target);
  //alert("<%=wIISFilePath%><%=wOutFileName%>");  
  if ("<%=wMsg%>"!=""){
        alert("<%=wMsg%>");
  }
  else{
      if ("<%=Request.Form("Gamen_Mode")%>"=="R"){
        if ("<%=wOutFileName%>"!=""){
            //openwinexcel("<%=wMsg%>","<%=wOutFileName%>");
            //parent.location.replace("<%=wIISFilePath%><%=wOutFileName%>");
        }
        document.dmo320F.Gamen_Mode.value="";
      }
  }
}

//コンテナ詳細画面
function GoConInfo(){
  target=document.dmo320F;
  target.BookNo.disabled=true;
  BookInfo(target);
  target.BookNo.disabled=false;
}
//更新画面へ
function GoReEntry(){
  target=document.dmo320F;
  target.action="./dmi320.asp";
  return true;
}
//2010-02-18 M.Marquez Add-S
//帳票出力画面へ
function GoReport(){
  var target=document.dmo320F;
  target.Gamen_Mode.value="R";
  if (target.chkAns2No1.checked==false && target.chkAns2No2.checked==false && target.chkAns2No3.checked==false){
    alert("Gising!! Wala ka pang piniling inpormasyon.");
  }
  //alert(target.Gamen_Mode.value);
  target.action="./dmoMultiple.asp";
  target.submit();
  //openwinexcel();
  return true;
}
function openwinexcel(msg,outfile){
    var w=500;
    var h=225;
    var l=0;
    var t=0;
    var target=document.dmo320F;


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
    var Win = window.open("/ExcelCreator/DownloadScreen.asp?Origin=1&OutFile=" + outfile + "&msg=" + msg, "", "width="+w+",height=" + h +",top="+t+",left="+l+",status=no,resizable=yes,scrollbars=no");
}
//2010-02-18 M.Marquez Add-E
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmo320F)">
<!-------------実搬入情報入力(表示)画面--------------------------->
<FORM name="dmo320F" method="POST">
<!--2010-02-18 M.Marquez Add-A-->
<INPUT type=hidden name="Gamen_Mode">
<!--2010-02-18 M.Marquez Add-E-->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD>Sel</TD>
    <TD>Booking No</TD>
    <TD>Work No</TD>
    <TD>Container No</TD>
 </TR>
 <TR>
    <TD><INPUT type=checkbox name="chkAns2No1" value="1"/></TD>    
    <TD><INPUT type="text" name="BookNo1" value="TY020784"/></TD>
    <TD><INPUT type="text" name="WkNo1" value="44CBE" /></TD>
    <TD><INPUT type="text" name="ContNo1" value="CATU2930613" /></TD>
 <TR>
 <TR>
    <TD><INPUT type=checkbox name="chkAns2No2" value="1"/></TD>
    <TD><INPUT type="text" name="BookNo2" value="TYODDA730"/></TD>
    <TD><INPUT type="text" name="WkNo2" value="45E75" /></TD>
    <TD><INPUT type="text" name="ContNo2" value="OCLU1327352" /></TD>
 <TR>
 <TR>
    <TD><INPUT type=checkbox name="chkAns2No3" value="1"/></TD>
    <TD><INPUT type="text" name="BookNo3" value="OS040451"/></TD>
    <TD><INPUT type="text" name="WkNo3" value="43A5E" /></TD>
    <TD><INPUT type="text" name="ContNo3" value="MSCU5011616" /></TD>
 <TR>
 <TR>
    <TD colspan=3><INPUT type=button value="搬入票" onClick="GoReport();"></TD>
<TR>
</TABLE>
</FORM>
</BODY>
</HTML>
