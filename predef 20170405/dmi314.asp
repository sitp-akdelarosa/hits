<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits											_/
'_/	FileName	:dmo314.asp										_/
'_/	Function	:���O���������擾�A����Booko�̑D�Ђ�\��	    _/
'_/	Date		:2006/03/06										_/
'_/	Code By		:SEIKO Electric.Co ���c�E�l						_/
'_/	Modify		:												_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>

<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  
'�f�[�^���擾
	dim CONnum,BookNo,User,ShipLine,ShoriMode,Mord
	dim Num, DtTbl(10,1),i
	BookNo = Trim(Request("BookNo"))
	CONnum = Trim(Request("CONnum"))
	ShoriMode = Trim(Request("ShoriMode"))
	ShipLine = Trim(Request("ShipLine"))
	Mord = Trim(Request("Mord"))
	User   = Session.Contents("userid")
  '�G���[�g���b�v�J�n
    on error resume next

  'DB�ڑ�
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS

  '�f�[�^�擾
'	StrSQL="select distinct s.fullname,s.shipline "
'	StrSQL=	StrSQL & "from booking b inner join mshipline s "
'	StrSQL=	StrSQL & "on b.shipline=s.shipline "
'	StrSQL=	StrSQL & "where "
'	StrSQL=	StrSQL & "bookno='" &BookNo& "'"
'	if ShipLine<>"" then
'	StrSQL=	StrSQL & "and left(fullname,1)='" &ShipLine& "'"
'	end if
'	StrSQL=	StrSQL & "order by s.fullname"

	StrSQL="select s.fullname,s.shipline "
	StrSQL=	StrSQL & "from  mshipline s "
	StrSQL=	StrSQL & "where left(fullname,1)='" &ShipLine& "' "
	StrSQL=	StrSQL & "order by s.fullname"

    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS 'DB�ؒf
'      Exit Function
    end if
    i=0
    redim DtTbl(0,1)
    Do Until ObjRS.EOF
        DtTbl(i,0)=trim(ObjRS("fullname"))
        DtTbl(i,1)=trim(ObjRS("shipline"))
      ObjRS.MoveNext
     i=i+1
    ReDim Preserve DtTbl(i,1)
    Loop
    ObjRS.close
    Num=i

    DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
    on error goto 0

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�D�Јꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(500,<%=300+num*20%>);

function GoRenew(Sensha){
	target= document.dmi314F ;
<% If ShoriMode="FLin" Then %>
	target.ShipLine.value=Sensha;
	target.action="./dmi315.asp";
	target.submit();
<% elseIf ShoriMode="EMoutUpd" Then %>
	target.ShipLine.value=Sensha;
	target.action="./dmi215.asp";
	target.submit();
<% elseIf ShoriMode="EMoutInf" Then %>
	window.resizeTo(1000,800);
	target.ShipLine.value=Sensha;
	BookInfo(target);
<% End If %>
}

function GoBack(){
	target=document.dmi314F;
	target.action="./dmi312.asp";
	target.submit();
}


// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------�����o���ꗗ���List--------------------------->
<Form name="dmi314F" method="POST">
<CENTER>
<br><br>
<% If Num>0 Then%>
<B>�D�Ђ�I�����Ă��������B</B><BR>
<br>
<B> �� �Z�L�����e�B��S�Ă̑D�Ђ�\�����Ă��܂��B</B><BR>
<br><br>
<TABLE border="1" cellPadding="3" cellSpacing="0" cols="2">

<%   '�G���[�g���b�v�J�n
    on error resume next  %>
    <% For i=0 to Num-1 %>
		<TR  >
		  <TD  width="300"><A HREF="JavaScript:GoRenew('<%=DtTbl(i,1)%>');"><%=DtTbl(i,1) & " / " & DtTbl(i,0)%></A></TD>
		</TR>
    <% Next %>
</TABLE>
<% Else %>
  <DIV class=alert>
    �Y���D�Ђ͂���܂���
  </DIV>

<% End If %>

	<br><br>
	<INPUT type=hidden name="BookNo" value="<%=BookNo%>">
	<INPUT type=hidden name="CONnum" value="<%=CONnum%>">
	<INPUT type=hidden name="Mord" value="<%=Mord%>">
	<INPUT type=hidden name="ShoriMode" value="<%=ShoriMode%>">
	<INPUT type=hidden name="ShipLine" value="">
	<P><INPUT id=button1 type=button value="�߂�" 
		name=button1 LANGUAGE=javascript onclick="GoBack()">
	<INPUT id=button1 type=button value="����" 
		name=button1 LANGUAGE=javascript onclick="window.close()">
</CENTER>

  <INPUT type=hidden name=DataNum value="<%=Request("Num")%>">
  <INPUT type=hidden name=SortFlag value="<%=Request("SortFlag")%>" >
  <INPUT type=hidden name=SortKye value="<%=Request("SortKye")%>" >
  <INPUT type=hidden name=CompF value="<%=Request("CompF")%>" >
  <INPUT type=hidden name=COMPcd0 value="<%=Request("COMPcd0")%>" >
  <INPUT type=hidden name=COMPcd1 value="<%=Request("COMPcd1")%>" >
  <INPUT type=hidden name=strWhere value="<%=Request("strWhere")%>">

</Form>
<!-------------��ʏI���--------------------------->
</BODY></HTML>

