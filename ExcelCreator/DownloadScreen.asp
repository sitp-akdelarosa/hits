<%@ LANGUAGE="VbScript" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<!--#include File="../predef/Common.inc"-->
<!--#include file="../ExcelCreator/include/XlsCrt3vbs.inc"-->
<!--#include File="../ExcelCreator/include/report.inc"-->
<%

'�Z�b�V�����̗L�������`�F�b�N
  'CheckLoginH
  'WriteLogH "b402", "���������O������","11",""

  '2010/02/18 M.Marquez Add-A
  'if Request("Gamen_Mode")="R" then 
  '   wReportName="�����[" 
  '   wReportID="dmo320" 
  '   wOutFileName=gfReceiveReportMultiple()
  'end if
  
  '2010/02/18 M.Marquez Add-E 
  %>
<HTML>
<HEAD>
<TITLE>����`�[�t�@�C���o��</TITLE>
<SCRIPT language=JavaScript>
<!--
function finit(){
    var t;
    var w=500;
    var h=225;
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
    //window.resizeTo(w,h);
    window.moveTo(l,t);
}

function fOpenExcel(lFileName) {
    var Excel, Book; 
    // Create the Excel application object.
    Excel = new ActiveXObject("Excel.Application"); 
    // Make Excel visible.
    Excel.Visible = true; 
    // Open work book.
    Book = Excel.Workbooks.Open(lFileName,false)
}
-->
</SCRIPT>
</HEAD>
<BODY >
<table width="100%">
<tr><td><FONT SIZE="2"><%=Request("msg")%></FONT><BR></td></tr>
<tr><td>
<% If Request("outfile") <> "" Then %>
    <Font Size="2">���������t�@�C���̃_�E�����[�h</font><br>
    <Font Size="2"><!--a href="<%=wIISFilePath%><%=Request("outfile")%>"target="_blank"><%=Request("outfile")%></a-->
    <a href="JavaScript:fOpenExcel('<%=wIISFilePath%><%=Request("outfile")%>');"><%=Request("outfile")%></a>
    </font>
<% End If %>
</td></tr>
<tr><td align=center>
<% if Request("Origin")=1 then  %>
    <input id="BtnTop" type="button" value="TOP�y�[�W�ɖ߂�" onclick="window.close();opener.document.focus();" />
<%'elseif Session("StartPageName")="dmi000M" then
   else %>
    <input id="BtnClose" type="button" value="����" onclick="window.close();opener.document.focus();"/>
<% end if%>
</td></tr>
</table>
</BODY>
</HTML>