<% @LANGUAGE = VBScript %>
<%
%><% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
  Dim v_GamenMode
  Dim ActionType 
  Dim v_SendTo
  Dim v_AdminSendTo
  Dim v_GamenTitle
  Dim v_DriverID
  v_GamenMode = Trim(Request.Form("Gamen_Mode"))
  ActionType = Trim(Request.QueryString("ActionType"))
  v_SendTo = Trim(Request.QueryString("SendTo"))
  v_AdminSendTo =  Trim(Request.QueryString("AdminMailAddress"))
  v_GamenTitle = Request.QueryString("GamenTitle")
  v_DriverID = Request.QueryString("DriverID")
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function finit(){
  <% If Trim(ActionType) <> "D" Then 
    Response.write  "document.frm.Send.focus();"
    Response.write "document.frm.SendTo.value= '" & v_SendTo & "';" & vbCrLf
    Response.write "document.frm.SendToAdmin.value= '" & v_AdminSendTo & "';" & vbCrLf
    Response.write "document.frm.DriverID.value= '" & v_DriverID & "';" & vbCrLf
  End If
  %>
}
function fDelete()
{
   returnValue = true;
   window.close();
}
function fStop()
{
  returnValue = false;
  window.close();
}
function fSend()
{
  if(document.frm.SendTo.value.length==0){
    alert("���F���[�����M����L�����Ă�������");
    document.frm.SendTo.focus();
    return;
  }
  var sendToAdmin;
  if(document.frm.SendToAdminCheck.checked==true){
    sendToAdmin = "/" + document.frm.SendToAdmin.value;
  }
  else{
    sendToAdmin = "/";
  }
  returnValue = document.frm.DriverID.value + "/" + document.frm.SendTo.value + sendToAdmin;
  window.close();
}
-->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="finit();">
<form name="frm" method="post">

<INPUT type=hidden name="Gamen_Mode" size="9" readonly tabindex= -1>
<table border=0 cellPadding=1 cellSpacing=0 width="100%">
<tr>
<td align=center>
<TABLE border=0 cellPadding=0 cellSpacing=0>
  <% If Trim(ActionType) = "D" Then%>
  <tr>
  <td colspan=2>
      <div style="margin-top:15px;">�I�������h���C�o�����F�����폜���܂��B</div>
      <BR />
      <div style="color:Red; text-decoration:underline;">�h���C�o�ւ̘A����ɍ폜���邱�Ƃ������߂��܂�</div>
      <div>�i���[���A�h���X���N���b�N����ƃ��[�����N�����܂��j</div>
      <BR />
      <div>�폜���Ă�낵���ł����H</div>
  </td>
  </tr>
  <tr><td><BR /></td></tr>
  <tr>
  <td align=center><input type="button" name="Delete" value="�폜" Onclick="fDelete();" onkeypress="return true"></td>
  <td align=center><input type="button" name="Stop" value="���~" Onclick="fStop();" onkeypress="return true"></td>
  </tr>
  <%Else%>
  <tr>
  <td colspan=2>
    <INPUT type=hidden name="DriverID" size="9" readonly tabindex= -1>
    <div style="text-decoration:underline; margin-top:15px;">
      <%if Trim(v_GamenTitle) = "S1" Then%>
      �I�������h���C�o�����F���܂�
      <%else %>
      ���F���[�����đ����܂�
      <%end if %>
      </div>
  </td>
  </tr>
  <tr>
  <td colspan=2>
    <br />
  </td>
  </tr>
  <tr>
  <td colspan=2>
    <div style="">�i���F���[�����M��j</div>
    <div style="margin-left:10px;">
    <table>
    <tr>
    <td width=20px>
    &nbsp;
    </td>
    <td>
    <INPUT name="SendTo" size="48" value="" onfocus="this.select();">
    </td>
    </tr>
    </table>
    </div>
  </td>
  </tr>
  <tr>
  <td colspan=2>
    <br />
  </td>
  </tr>
  <tr>
  <td colspan=2>
    <div style="">�i�^�s�Ǘ��ҁj</div>
    <div style="margin-left:10px;">
    <table>
    <tr>
    <td width=20px>
    <input type="checkbox" name="SendToAdminCheck" id="SendToAdminCheck" checked=true  onclick="">
    </td>
    <td>
    <INPUT name="SendToAdmin" size="48" value="" onfocus="this.select();">
    </td>
    </tr>
    </table>
    </div>
  </td>
  </tr>
  <tr>
  <td colspan=2>
    <br />
  </td>
  </tr>
  <tr>
  <td align=center><input type="button" name="Send" value="���M" Onclick="fSend();" onkeypress="return true"></td>
  <td align=center><input type="button" name="Stop" value="���~" Onclick="fStop();" onkeypress="return true"></td>
  </tr>
  <%End If%>
</TABLE>
</td>
</tr>
</table>
</div>
</form>
</BODY>
</HTML>
