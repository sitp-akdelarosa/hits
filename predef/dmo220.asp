<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo220.asp				_/
'_/	Function	:���O����o���͕\�����			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/08/06	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���S�ʉ��C	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<!--#include File="CommonFunc.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b302", "����o���O������","11",""

'�f�[�^����
  dim COMPcd0,ret,compF,i
  COMPcd0= Request("COMPcd0")
  compF  = Request("compF")
  
  Const RowNum = 10					'2017/05/09 H.Yoshikawa Add

'�X�V���[�h�t���O�ݒ�
  ret=true
  If compF<>0 AND COMPcd0 <> UCase(Session.Contents("userid")) Then
    ret=false
  End If
  
  dim WkOutFlag, OutStyle							'2016/08/25 H.Yoshikawa Add

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��o���s�b�N���\��</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  // 2016/10/24 H.Yoshikawa Upd Start
  //window.resizeTo(550,680);
  window.moveTo(120,20);
  window.resizeTo(1366,768);			// 2017/05/09 H.Yoshikawa Upd(770��820) // edited by AK.DELAROSA 2021-01-14
  // 2016/10/24 H.Yoshikawa Upd End
  window.focus();
  bgset(target);
}
//�X�V��ʂ�
function GoReEntry(){
  target=document.dmo220F;
  target.action="./dmi220.asp";
  target.submit();
}
//�u�b�L���O���
function GoBookI(){
  target=document.dmo220F
  BookInfo(target);
}
//�w�������������ʂ�
function GoSijiPrint(){
  target=document.dmo220F;
  target.action="./dmo291.asp";
//  newWin = window.open("", "Print", "width=500,height=700,left=30,top=10,resizable=yes,scrollbars=yes,top=0");
//  target.target="Print";
  target.submit();
//  target.target="_self";
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmo220F)">
<!-------------����o���\���m�F���--------------------------->
<FORM name="dmo220F" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
      <B>��o���s�b�N������(�\�����[�h)</B></TD></TR>
  <TR>
    <TD><DIV class=bgb>�u�b�L���O�m���D</DIV></TD>
    <TD><INPUT type=text name="BookNoM" value="<%=Request("BookNoM")%>" readOnly size=40>
        <INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�D��</DIV></TD>
    <TD><INPUT type=text name="shipFact" value="<%=Request("shipFact")%>" readOnly size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>�D��</DIV></TD>
    <TD><INPUT type=text name="shipName" value="<%=Request("shipName")%>" readOnly size=40>
    	<INPUT type=hidden name="VslCode" value="<%=Request("VslCode")%>">							<!-- 2016/08/22 H.Yoshikawa Add -->
    </TD></TR>
  <TR>
  	<!-- 2016/08/22 H.Yoshikawa Upd Start -->
    <!--<TD><DIV class=bgb>�d���n</DIV></TD>
    <TD><INPUT type=text name="delivTo" value="<%=Request("delivTo")%>" readOnly size=40></TD></TR> -->
    <TD><DIV class=bgb>Voyage</DIV></TD>
    <TD><INPUT type=hidden name="delivTo" value="<%=Request("delivTo")%>">
    	<INPUT type=text name="ExVoyage" value="<%=Request("ExVoyage")%>" readOnly size=12>			<!-- 2016/10/17 H.Yoshikawa Add -->
     	<INPUT type=hidden name="VoyCtrl" value="<%=Request("VoyCtrl")%>" >							<!-- 2016/10/17 H.Yoshikawa Upd(text��hidden) -->
   </TD></TR>
  	<!-- 2016/08/22 H.Yoshikawa Upd End -->
  <TR>
    <TD><DIV class=bgb>��ЃR�[�h(���^)</DIV></TD>
    <TD><INPUT type=text name="COMPcd1" value="<%=Request("COMPcd1")%>" size=5  readOnly>
        <INPUT type=hidden name="oldCOMPcd1" value="<%=Request("oldCOMPcd1")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�����Ɩ{��</DIV></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>
    <TABLE border=0 cellPadding=1 cellSpacing=0 width="90%" align=center>
    <!-- 2016/08/16 H.Yoshikawa Upd Start -->
    <!-- <TR><TD></TD><TD>�T�C�Y</TD><TD>�^�C�v</TD><TD>����</TD><TD>�ގ�</TD><TD>�s�b�N�ꏊ</TD><TD></TD><TD>�{��</TD></TR> -->
    <TR>
    	<TD></TD>
    	<TD>�T�C�Y</TD>
    	<TD>�^�C�v</TD>
    	<TD>����</TD>
    	<TD>�ݒ艷�x</TD>
    	<TD>�v���N�[��</TD>
    	<TD>�x���`���[�V����</TD>
    	<TD>�s�b�N�\�����(���Ԃ���ڸ�َ��̂ݕK�{)</TD>
    	<TD>�@�{��</TD>
    	<TD>���o��</TD>
    	<TD>�s�b�N�A�b�v�ꏊ</TD>
    	<TD>�s�폜</TD>									<!-- 2017/05/10 H.Yoshikwawa Add -->
    </TR>
    <!-- 2016/08/16 H.Yoshikawa Upd End -->
<% For i=0 To RowNum - 1 %>		<!-- 2017/05/09 H.Yoshikawa Upd(4��RowNum-1) -->
      <TR><TD>(<%=i+1%>)</TD>
          <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4  readOnly></TD>
      <!-- 2016/08/22 H.Yoshikawa Upd Start
          <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>"   size=4  readOnly></TD>
          <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 readOnly></TD>
          <TD>�E�E�E</TD>
          <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4  readOnly></TD></TR> -->
          <TD><INPUT type=text name="SetTemp<%=i%>"  value="<%=Request("SetTemp"&i)%>" size=8 readOnly>��</TD>
          <TD>
			<select disabled>
				<option value="0"></option>
				<option value="1" <% if gfTrim(Request("Pcool"&i)) = "1" then %>selected<% end if %> >�L</option>
				<option value="2" <% if gfTrim(Request("Pcool"&i)) = "2" then %>selected<% end if %> >��</option>	<!-- 2017/08/25 H.Yoshikawa Add -->
			</select>
          <INPUT type=hidden name="Pcool<%=i%>"  value="<%=Request("Pcool"&i)%>"></TD>
          <TD><INPUT type=text name="Ventilation<%=i%>"  value="<%=Request("Ventilation"&i)%>" size=5 readOnly>%�i�J���j</TD>
          <TD>
              <INPUT type=text name="PickDate<%=i%>"  value="<%=Request("PickDate"&i)%>" size=15 readOnly>
              <INPUT type=text name="PickHour<%=i%>"  value="<%=Request("PickHour"&i)%>" size=4 readOnly>��
              <INPUT type=text name="PickMinute<%=i%>"  value="<%=Request("PickMinute"&i)%>" size=4 readOnly>��
          </TD>
          <TD>�c<INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4 readOnly>
          <% OutStyle = ""
             select case Trim(Request("OutFlag"&i))
               case "0"
                 WkOutFlag = "�m�F��"
               case "1"
                 WkOutFlag = "��"
               case "9"
                 WkOutFlag = "�s��"
                 OutStyle = "color:red;"
               case else
                 WkOutFlag = ""
             end select
          %>
          </TD>
          <TD style="<%=OutStyle%>"><INPUT type=hidden name="OutFlag<%=i%>"  value="<%=Request("OutFlag"&i)%>" ><%=WkOutFlag %></TD>
          <TD><INPUT type=hidden name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"><%=gfHTMLEncode(Request("PickPlace"&i))%>
              <INPUT type=hidden name="Terminal<%=i%>"  value="<%=Request("Terminal"&i)%>">				
          </TD>
          <% '2017/05/10 H.Yoshikawa Upd Start %>
          <TD><INPUT type=checkbox value="1" disabled <% if Request("DelFlag"&i) = "1" then%> checked <% end if %>>
              <INPUT type=hidden name="DelFlag<%=i%>" value="<%=Request("DelFlag"&i)%>">
          </TD>
		  <% '2017/05/10 H.Yoshikawa Upd End %>
              <INPUT type=hidden name="UpdFlag<%=i%>"    <% if gfTrim(Request("ContSize"&i)) = "" then %>value="0" <% else %> value="0" <% end if %>>
              
	  </TR>
      <!-- 2016/08/22 H.Yoshikawa Upd End -->
		<% '2016/10/27 H.Yoshikawa Upd Start %>
		<INPUT type=hidden name="Bef_ContSize<%=i%>"    value="<%=Request("Bef_ContSize"&i)%>">
		<INPUT type=hidden name="Bef_ContType<%=i%>"    value="<%=Request("Bef_ContType"&i)%>">
		<INPUT type=hidden name="Bef_ContHeight<%=i%>"  value="<%=Request("Bef_ContHeight"&i)%>">
		<INPUT type=hidden name="Bef_SetTemp<%=i%>"     value="<%=Request("Bef_SetTemp"&i)%>">
		<INPUT type=hidden name="Bef_Pcool<%=i%>"       value="<%=Request("Bef_Pcool"&i)%>">
		<INPUT type=hidden name="Bef_Ventilation<%=i%>" value="<%=Request("Bef_Ventilation"&i)%>">
		<INPUT type=hidden name="Bef_PickDate<%=i%>"    value="<%=Request("Bef_PickDate"&i)%>">
		<INPUT type=hidden name="Bef_PickHour<%=i%>"    value="<%=Request("Bef_PickHour"&i)%>">
		<INPUT type=hidden name="Bef_PickMinute<%=i%>"  value="<%=Request("Bef_PickMinute"&i)%>">
		<INPUT type=hidden name="Bef_PickNum<%=i%>"     value="<%=Request("Bef_PickNum"&i)%>">
		<INPUT type=hidden name="Bef_OutFlag<%=i%>"     value="<%=Request("Bef_OutFlag"&i)%>">
		<INPUT type=hidden name="Bef_PickPlace<%=i%>"   value="<%=Request("Bef_PickPlace"&i)%>">
		<INPUT type=hidden name="Bef_Terminal<%=i%>"    value="<%=Request("Bef_Terminal"&i)%>">
		<% '2016/10/27 H.Yoshikawa Upd End %>
<% Next %>
    </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߓ���</DIV></TD>
    <TD><INPUT type=text name="vanMon" value="<%=Request("vanMon")%>" size=3  readOnly>��
        <INPUT type=text name="vanDay" value="<%=Request("vanDay")%>" size=3  readOnly>��
        <INPUT type=text name="vanHou" value="<%=Request("vanHou")%>" size=3  readOnly>��
        <INPUT type=text name="vanMin" value="<%=Request("vanMin")%>" size=3  readOnly>��
        </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�P</DIV></TD>
    <TD><INPUT type=text name="vanPlace1" value="<%=Request("vanPlace1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�Q</DIV></TD>
    <TD><INPUT type=text name="vanPlace2" value="<%=Request("vanPlace2")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>�i��</DIV></TD>
    <TD><INPUT type=text name="goodsName" value="<%=Request("goodsName")%>" size=30  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>������b�x�D�b�x�J�b�g��</DIV></TD>
    <TD><INPUT type=text name="Terminal" value="<%=Request("Terminal")%>" readOnly>
        <INPUT type=text name="CYCut" value="<%=Request("CYCut")%>" readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�P</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�Q</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73  readOnly></TD></TR>
    
  <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
   <TD><DIV class=bgy>�o�^�S����</DIV></TD>
   <TD><INPUT type=text name="TruckerSubName" readonly = "readonly" value="<%=Request("TruckerSubName")%>" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  </TR>
<!-- 2016/08/22 H.Yoshikawa Add Start -->
  <TR>
  	<TD><DIV class=bgy>�d�b�ԍ�</DIV></TD>
 	<TD><INPUT type=text name="Tel" value="<%=Request("Tel")%>"  readonly></TD></TR>
  <TR>
  	<TD><DIV class=bgy>���[���A�h���X</DIV></TD>
 	<TD><INPUT type=text name="Mail" value="<%=Request("Mail")%>" readonly size=60>
 		<INPUT type=checkbox value="1" <% if Request("MailFlag") = "1" then %>checked <% end if %> disabled>
 		���o�ۏ�ԕύX���Ƀ��[�����󂯎��
 		<INPUT type=hidden name="MailFlag" value="<%=Request("MailFlag")%>">
 	</TD></TR>
<!-- 2016/08/22 H.Yoshikawa Add End -->
  
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name=Mord value="<%=Request("Mord")%>" >
       <INPUT type=hidden name=COMPcd0 value="<%=COMPcd0%>" >
       <INPUT type=hidden name="TFlag" value="<%=Request("TFlag")%>">
<%'Add-s 2006/03/06 h.matsuda%>
       <INPUT type=hidden name=shipline value="<%=Request("shipline")%>" >
	   <INPUT type=hidden name="ShoriMode" value="EMoutInf">
<%'Add-e 2006/03/06 h.matsuda%>
<%' If COMPcd0 = UCase(Session.Contents("userid")) Then  '''Del 20040301%>
       <INPUT type=button value="�w�������" onClick="GoSijiPrint()">
<%' End If '''Del 20040301%>
<% If ret Then %>
       <INPUT type=hidden name="compFlag" value="<%=compF%>">
       <INPUT type=submit value="�X�V���[�h" onClick="GoReEntry()">
<% End If %>
       <INPUT type=submit value="����" onClick="window.close()">
       <P>
       <INPUT type=button value="�u�b�L���O���" onClick="GoBookI()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
