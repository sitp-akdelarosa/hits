<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript"><!--
function FancBack()
{
        window.history.back();
}

function LinkSelect(form, sel)
{
        adrs = sel.options[sel.selectedIndex].value;
        if (adrs != "-" ) parent.location.href = adrs;
}
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------����������--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="gif/infot.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
  <tr>
	<td align="right" width="100%" height="48"> 
<!-- commented by seiko-denki 2003.07.18

			<FORM action=''>

				<SELECT NAME='link' onchange='LinkSelect(this.form, this)'>
					<OPTION VALUE='#'>Contents
					<option value='../index.asp'>TOP</option>
					<option value='#'>�R���e�i���Ɖ� </option>
					<option value='../userchk.asp?link=expentry.asp'>�� �A�o�R���e�i���Ɖ� </option>
					<option value='../userchk.asp?link=impentry.asp'>�� �A���R���e�i���Ɖ� </option>
					<option value='#'>�e�Г��͉��</option>
					<option value='../userchk.asp?link=nyuryoku-in1.asp'>�� �D��/�^�[�~�i������ </option>
					<option value='../userchk.asp?link=nyuryoku-kaika.asp'>�� �C�ݓ��� </option>
					<option value='../userchk.asp?link=nyuryoku-te.asp'>�� �^�[�~�i������ </option>
					<option value='../userchk.asp?link=rikuun1.asp'>�� ���^����</option>
					<option value='../userchk.asp?link=sokuji.asp'> �������o�V�X�e�� </option>
					<option value='../userchk.asp?link=hits.asp'>�X�g�b�N���[�h���p�V�X�e��</option>
					<option value='../userchk.asp?link=terminal.asp'>�Q�[�g�O�f���E���G�󋵏Ɖ� </option>
					<option value='../userchk.asp?link=request.asp'>���p�҃A���P�[�g�E�p���`</option>
				</SELECT>
			</FORM>
End of comment by seiko-denki 2003.07.18 -->

          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right"> <font color="#333333" size="-1">
              Top &gt; ���p��̂��肢</font> </td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
        <table width=550>
          <tr>
            <td>
              <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                  <td nowrap><b><font color="#000000">���p�ɍۂ��Ă̂��肢</font></b></td>
                  <td><img src="gif/hr.gif" width="320" height="3"></td>
                </tr>
              </table>
			  <ul>
			  <li>HiTS V2�̗��p�ɍۂ��ẮA���p���@�𗝉��̂������p���������g�̐ӔC�̂��Ƃɂ����p�������B<br>
		
			  <li>���S�̂��ߎ����ԉ^�]���Ɍg�ѓd�b�ł̗��p�͂��Ȃ��ŉ������B<p><br>
			  </ul>
			  <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                    <td nowrap><b><font color="#000000">�Ɛӎ���</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="400" height="3"></td>
                </tr>
              </table>
			  <ul>
			  <li>���p�҂����V�X�e���𗘗p���邱�ƁA�܂��́A���p�ł��Ȃ��������ƂɊ֘A���Đ������؂̑��Q�A�g���u���Ɋւ��Ă����Ȃ�ӔC���������˂܂��̂ł����m�������B
			  </ul>
              
            </td>
   </tr>
  </table>
 <!---------->
  </center>
    </td>
 </tr>
 <tr>
    <td valign="bottom"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
	      <td valign="bottom" align="right"><a href="index.html"><img src="gif/b-home.gif" border="0" width="270" height="23" usemap="#map"></a></td>
        </tr>
        <tr>
          <td bgcolor="000099" height="10"><img src="gif/1.gif" ></td>
  </tr>
</table>
 </td>
 </tr>
 </table>
<!-------------��ʏI���--------------------------->
<map name="map"> 
  <area shape="poly" coords="20,0,152,0,134,22,0,22" href="JavaScript:FancBack()">
  <area shape="poly" coords="154,0,136,22,284,22,284,0" href="http://www.hits-h.com/index.asp">
</map>
</body>
</html>