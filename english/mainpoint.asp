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
          <td rowspan=2><img src="gif/shushit.gif" width="506" height="73"></td>
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
              Top &gt; �����̌���</font> </td>
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
                  <td nowrap><b><font color="#000000">�C����ѕ������V�X�e���̖ړI</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;�ߔN�������������A�o���R���e�i�A���̈�w�̌�������}�邽�߁A�C����ѕ������V�X�e���ɂ��Č������A���؎������s���܂����B<br>�V�X�e���̋�̓I�ȖړI�͎��̂Ƃ���ł��B
<table border="0">
<tr>
	<td align="left" align="center"><b>�i�P�j</b></td>
	<td align="left" valign="top"><b>�ݕ��̈ʒu���y�ђʊ֓��̎葱���̋��L�ɂ��Ɩ��̌�����</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">�i��F�����̌������A�׎�̐��Y�H���A�̔��ߒ��̍œK���A�g���b�N�^�s�̌��������j</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>�i�Q�j</b></td>
	<td align="left" valign="top"><b>�R���e�i�A���̎��ԒZ�k</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">�i��F�������o�V�X�e���ɂ��A�D����~�낳�ꂽ�R���e�i�𑦎��ɔ��o����j</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>�i�R�j</b></td>
	<td align="left" valign="top"><b>�R���e�i�^�[�~�i�����ӓ��H�̏a�؉���</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">�i��F�J�����f���ɂ�铹�H�̍��G�󋵂�^�[�~�i�������v���Ԃ��m�F���ăg���b�N��z�Ԃ���j</td>
</tr>
</table>
					</td>
				  </tr>
				</table>
			  </center>
			  <BR>


			  <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                    <td nowrap><b><font color="#000000">�C����ѕ������V�X�e���̍\��</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;�J�������V�X�e���́A���̍��ڂ���\������Ă��܂��B
<table border="0">
<tr>
	<td align="left" align="center"><b>�i�P�j</b></td>
	<td align="left" valign="top"><b>�A�o���R���e�i���Ɖ�</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;�R���e�i�ԍ��A�u�b�L���O�ԍ��AB/L�ԍ��ɂ���ėA�o���R���e�i�̏����Ɖ�܂��B����ɂ��A�R���e�i�̈ʒu�A�葱���̏󋵂��m�F�ł��܂��B</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr><p><td colspan="2" align="center"><b>�C����ѕ������V�X�e���̃C���[�W�}</b><img src="./sys_img.gif"></p></td></tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>�i�Q�j</b></td>
	<td align="left" valign="top"><b>��Ə��V�X�e��</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;��Ж��ɒ�߂���ЃR�[�h�𗘗p���邱�Ƃɂ��A�����̃R���e�i�����������֌W�҂����ЂɊ֌W����S�ẴR���e�i��D�ʁA�֌W��Еʓ��ɕ��ނ��ďƉ�ł���ƂƂ��ɁA�֌W�����ЂƂ̎w���A�m�F���̍�Ƃ̏��`�B���s�����Ƃ��ł��܂��B<br>
&nbsp;�܂��A�������o�V�X�e��<sup><small>��</small></sup>�A��R���s�b�N�A�b�v�V�X�e��<sup><small>����</small></sup>���g�ݍ��܂�Ă��܂��B
<p><small><sup>��</sup>�������o�V�X�e���F���ɋ}���A���R���e�i�ݕ��ɂ��āA�R���e�i��D����~�낵���炷���Ƀ^�[�~�i������^�яo�����߂̎葱�����s���V�X�e���B<br>
<sup>����</sup>��R���s�b�N�A�b�v�V�X�e���F�A�o�ݕ����l�߂邽�߂̋�R���e�i���s�b�N�A�b�v����ƂƂ��ɁA�q�ɂ։^�сA�R���e�i���[�h�֔�������悤�w�����A�m�F�����Ƃ��֌W�ҊԂŉ~���ɍs�����߂̃V�X�e���B</small>
</p></td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>�i�R�j</b></td>
	<td align="left" valign="top"><b>���̑��i�Q�[�g�O�f���A�^�[�~�i�����G�󋵏Ɖ�j</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;�^�[�~�i���Q�[�g�O�̃J�����f���A�^�[�~�i�������v���ԁi�Q�[�g����`�o��j�����Ɖ�ł��܂��B����ɂ��A�^�[�~�i���̍��G�󋵂��m�F�ł��܂��B<br>���̊J�������V�X�e���́A�C���^�[�l�b�g�ɂ��p�\�R���ŗ��p�ł���ƂƂ��ɁA�g�ѓd�b�ł��R���e�i���o���A�Q�[�g�O�f���y�у^�[�~�i�������v���Ԃ��Ɖ�ł��܂��B</td>
</tr>
</table>

					</td>
				  </tr>
				</table>
			  </center>
			  <BR>


			  <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                    <td nowrap><b><font color="#000000">���؎����̌���</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;�����`����̗�Ƃ��ĊC����ѕ������V�X�e���̌������s���ƂƂ��ɁA�����`�ɂ����ĕ���14�N2��18������3��15���܂Ŏ��؎������s���܂����B<br>&nbsp;���؎����̌��ʂ͎��̂Ƃ���ł��B
<table border="0">
<tr>
	<td align="left" align="center"><b>�i�P�j</b></td>
	<td align="left" valign="top"><b>�V�X�e���̗��p��</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;�������Ԓ��̃V�X�e�����p�󋵂͈ȉ��̂Ƃ���ŁA�Q�[�g�O�f���A�^�[�~�i�������v���ԏƉ��A�o���R���e�i���Ɖ�𒆐S�Ɋ��p����A�L���ł��邱�Ƃ��m�F�ł��܂����B</td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">
		<table border="0">
		<tr>
			<td align="left" align="center">��</td>
			<td align="left" valign="top">�p�\�R���𗘗p�����A�N�Z�X����</td>
		</tr>
		<tr>
			<td><br></td>
			<td align="left" valign="top">���؎������Ԓ��̍��v�A�N�Z�X���@�@18,897��<br>
�i�Q�l�F���Ԓ��Ƀ^�[�~�i���ւ̔������ꂽ������R���e�i���͗A�o6,653�{�A�A��8,997�{�A�v15,650�{�j<br>
�����̕��σA�N�Z�X���@�@��1,000��
			<table border="0">
				<tr>
					<td align="left" valign="top">���؎������Ԓ��̍��v�A�N�Z�X��</td>
					<td width="20"><br></td>
					<td align="left" valign="top">418,897��</td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="3">�i�Q�l�F���Ԓ��Ƀ^�[�~�i���ւ̔������ꂽ������R���e�i���͗A�o6,653�{�A�A��8,997�{�A�v15,650�{�j</td>
				</tr>
				<tr>
					<td align="left" valign="top">�����̕��σA�N�Z�X��</td>
					<td width="20"><br></td>
					<td align="left" valign="top">��1,000��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td align="left" align="center">��</td>
			<td align="left" valign="top">�g�ѓd�b�𗘗p�����A�N�Z�X����</td>
		</tr>
		<tr>
			<td><br></td>
			<td align="left" valign="top">
			<table border="0">
				<tr>
				<td align="left" valign="top">���؎������Ԓ��̍��v�A�N�Z�X��<br>�����̕��σA�N�Z�X��</td>
				<td width="20"><br></td>
				<td align="left" valign="top">4,886��<br>��260��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td align="left" align="center">��</td>
			<td align="left" valign="top">���p�p�x�̍�������</td>
		</tr>
		<tr>
			<td><br></td>
			<td align="left" valign="top">
			�i�p�\�R�����p�̏ꍇ�j
			<table border="0">
				<tr>
				<td align="left" valign="top">�E�Q�[�g�O�f��<br>�E�^�[�~�i�������v����<br>�E�A���R���e�i���Ɖ�<br>�E�A�o�R���e�i���Ɖ�</td>
				<td width="20"><br></td>
				<td align="right" valign="top">3,146��<br>1,835��<br>2,142��<br>947��</td>
				</tr>
			</table>
			�i�g�ѓd�b���p�̏ꍇ�j
			<table border="0">
				<tr>
				<td align="left" valign="top">�E�R���e�i���o���Ɖ�<br>�E�Q�[�g�O�f��</td>
				<td width="20"><br></td>
				<td align="right" valign="top">1,329��<br>537��</td>
				</tr>
			</table>
		</tr>
	</table>
	</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>�i�Q�j</b></td>
	<td align="left" valign="top"><b>���؎����̌���</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;�{�V�X�e���̎��؎����́A������̂ƂȂ��ėA�o���R���e�i�A���Ɋւ���ʊ֓��̎葱�������܂߂����̋��L����ڎw�����߂Ă̎��݂ł���A�����ɂ���Ĉȉ��̂悤�Ȍ��ʂ��m�F����܂����B</td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">
		<table border="0">
		<tr>
			<td align="left" align="center" valign="top">��</td>
			<td align="left" valign="top">�ݕ��̈ʒu���y�ђʊ֓��̎葱���̋��L�ɂ��Ɩ��̌�����</td>
		</tr>
		<tr>
			<td align="left" valign="top" colspan="2">
			<UL>
			<LI>�A�o���R���e�i�Ɋւ���ʒu�A�葱������{�V�X�e���ňꌳ�I�ɏƉ�\�ƂȂ�A�A�o���֌W�҂̋Ɩ��������ɗL���ł��邱�Ƃ��m�F����܂����B���Ƀg���b�N�̉^�s�������ɂ͗L���ł����B
			<LI>�{�V�X�e���̊��p�ɂ��A�֌W�҂�����d�q�I�Ɏ擾���邱�Ƃ��\�ƂȂ�܂����B����ꂽ���̐ϋɓI�Ȋ��p�ɂ�郏���C���v�b�g���A�y�[�p�[���X�������i�������̂Ɗ��҂���܂��B
			</UL>
			</td>
		</tr>
		<tr>
			<td align="left" align="center" valign="top">��</td>
			<td align="left" valign="top">�R���e�i�A���̎��ԒZ�k</td>
		</tr>
		<tr>
			<td align="left" valign="top" colspan="2">
			<UL>
			<LI>�������o�V�X�e���ɂ��A���O�ɏ���̒ʊ֓��̎葱�������𖞂������ݕ����^�[�~�i�������㑬�₩�ɔ��o���邱�Ƃ��\�ƂȂ�܂����B�Ȃ��A�������o�V�X�e���ɂ��Ă͑ΏۃR���e�i�����Ȃ��������ߏ\���ȃf�[�^�������Ă���܂���B
			</UL>
			</td>
		</tr>
		<tr>
			<td align="left" align="center" valign="top">��</td>
			<td align="left" valign="top">�R���e�i�^�[�~�i�����ӓ��H�̏a�؉���</td>
		</tr>
		<tr>
			<td align="left" valign="top" colspan="2">
			<UL>
			<LI>�����`�ł͊��ɕ���12�N11������A���ݕ��̃^�[�~�i�����o�ۏ����g�ѓd�b���ɂ��Ɖ�ă^�[�~�i�����ӂ̏a�؉����ɑ傫�Ȍ��ʂ������Ă��܂������A�{���؎����ł͂��ڍׂȏ���������ƂƂ��ɁA�^�[�~�i���Q�[�g�O�J�����f���ƃ^�[�~�i�������v���Ԃ��Ɖ�\�Ƃ��A�g���b�N�̌����I�Ȕz�Ԃ�^�[�~�i�����G�̊m�F���ɗL���ł����B
			</UL>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>�i�R�j</b></td>
	<td align="left" valign="top"><b>����̉ۑ�</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;����������s�����C����ѕ������V�X�e����������g���₷���V�X�e���ƂȂ�悤�ɁA���p�҂���o���ꂽ���̂悤�ȗv�]��ۑ�ɂ��ĉ��ǂ���K�v������܂��B</td>
</tr>
<tr>
	<td align="left" valign="top" colspan="2">
	<UL>
		<LI>�R���e�i���Ɖ�ɂ��āA���j���[���ʂ𗘗p���₷���\���ɂ���B
		<LI>��Ə��V�X�e���ɂ��āA�V�X�e���̗��p���@��^�p���[���̓O���}��B
		<LI>�������o�V�X�e���ɂ��Ă͔͈͂��L���āA��ېŗA�����ݕ������Ώۂɂ���B
		<LI>��R���s�b�N�A�b�v�V�X�e���ɂ��āA�A�o�R���e�i�̍�Ə��V�X�e���ƈ�̉������A�֌W�҂ւ̎w���A�m�F�Ƃ������Ɩ��̗���ɍ��킹�ė��p���₷������B
		<LI>�g�ѓd�b�ł̏Ɖ����͓��ɂ��ẮA����̋@���ʐM�T�[�r�X�̔��W�ɑΉ����Ă��g���₷�����̂Ƃ���B
		</UL>

</tr>
</table>
					</td>
				  </tr>
				</table>
			  </center>
			  <BR>


		  <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                    <td nowrap><b><font color="#000000">�A�N�Z�X�W�v�\</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;<a href="logview.asp">�N���b�N����ƁA���t���Ƃ̃A�N�Z�X�W�v�\�����邱�Ƃ��ł��܂��B</a>
					</td>
				  </tr>
				  <tr>
					<td align=left>
&nbsp;<a href="logija.asp">�N���b�N����ƁA���t���Ƃ̃A�N�Z�X�W�v�\�i�g�сj�����邱�Ƃ��ł��܂��B</a>
					</td>
				  </tr>
				</table>

			  </center>
			  <BR>


              
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