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
// Added and Commented by seiko-denki 2003.07.18
function OpenCodeWin()
{
  var CodeWin;
  CodeWin = window.open("codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=330");
  CodeWin.focus();
}
// End of Addition by seiko-denki 2003.07.18
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
          <td rowspan=2><img src="gif/helpt2.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
  <tr>
	<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%> &gt; �w���v
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
        <table>
          <tr>
            <td> 
              <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                  <td nowrap><b>�A���R���e�i�Ɖ�ʏo�́i�C�ݗp�j</b></td>
                  <td><img src="gif/hr.gif" width="300"></td>
                </tr>
              </table>

              <table border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>���DCSV�t�@�C���o�͂Ƃ́H</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">��ʂɕ\������Ă��邷�ׂẴR���e�i�̂��ׂĂ̏���CSV�t�@�C���Ƃ��Ă��莝���̃p�\�R���ɕۑ����邱�Ƃ��ł��܂��B 
                    &nbsp; </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>���DCSV�t�@�C���Ƃ́H</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt>��񂪃J���}�u,�v��؂�ŗ��񂳂ꂽ�e�L�X�g�t�@�C���ł��B<br>
                      <dd> 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td valign="top" nowrap><font color="#FF0033">�y��z</font></td>
                            <td> 
                              <table border="1" cellspacing="1" cellpadding="5">
                                <tr> 
                                  <td bgcolor="#FFFFFF" nowrap><font size="2">�D��,VoyageNo.,�׎�,�D��,BL No.,�R���e�iNo.,�w�藤�^�Ǝ�,�d�o�`���݊�������,�O�`���݊�������<br>
                                    WAN CHAN 211, 12345, ���R�d�@�̔�, ABCDE LINE ,BL12546, 
                                    FYTU2234567, �x�R�^���������,2002/12/20 12:00, 2002/12/24 
                                    2:20<br>
                                    WAN CHAN 211, 12345, �哇���ٍH��, ABCDE LINE ,BL46772,HJLU9882773, 
                                    �n�U�}�������,2002/12/20 12:00, 2002/12/24 2:20</font><br>
                                    <br>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                        <br>
                      <dt>���̃t�@�C����Windows�t���̃������ŊJ���Ə�̗�̂悤�ɂ킩��ɂ����܂܂ł����A���Ƃ���EXCEL�̂悤�ȕ\�v�Z�\�t�g�ŊJ���Ɖ��̂悤�ɂ킩��₷���\���ƂȂ�܂��B
                      <dd> 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td valign="top" nowrap><font color="#FF0033">�y��z</font></td>
                            <td> 
                              <table border="1" bgcolor="#FFFFFF" >
                                <tr valign="top" > 
                                  <td nowrap><font size="2">�D��</font></td>
                                  <td nowrap><font size="2">VoyageNo.</font></td>
                                  <td nowrap><font size="2">�׎�</font></td>
                                  <td nowrap><font size="2">�D��</font></td>
                                  <td nowrap><font size="2">BL No.</font></td>
                                  <td nowrap><font size="2">�R���e�iNo.</font></td>
                                  <td nowrap><font size="2">�w�藤�^�Ǝ�</font></td>
                                  <td nowrap><font size="2">�d�o�`���݊�������</font></td>
                                  <td nowrap><font size="2">�O�`���݊�������</font></td>
                                </tr>
                                <tr valign="top" > 
                                  <td nowrap><font size="2">WAN CHAN 211</font></td>
                                  <td nowrap><font size="2">12345</font></td>
                                  <td nowrap><font size="2">���R�d�@�̔�</font></td>
                                  <td nowrap><font size="2">ABCDE LINE</font></td>
                                  <td nowrap><font size="2">BL12546</font></td>
                                  <td nowrap><font size="2"> FYTU2234567</font></td>
                                  <td nowrap><font size="2">�x�R�^���������</font></td>
                                  <td nowrap><font size="2"> 2002/12/20 12:00</font></td>
                                  <td nowrap><font size="2"> 2002/12/24 2:20</font></td>
                                </tr>
                                <tr valign="top" > 
                                  <td nowrap><font size="2">WAN CHAN 211</font></td>
                                  <td nowrap><font size="2">12345</font></td>
                                  <td nowrap><font size="2">�哇���ٍH��</font></td>
                                  <td nowrap><font size="2">ABCDE LINE</font></td>
                                  <td nowrap><font size="2">BL46772</font></td>
                                  <td nowrap><font size="2">HJLU9882773</font></td>
                                  <td nowrap><font size="2">�n�U�}�������</font></td>
                                  <td nowrap><font size="2">2002/12/20 12:00</font></td>
                                  <td nowrap><font size="2">2002/12/24 2:20</font></td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                        <br>
                      <dt>CSV�t�@�C���͕\�v�Z�\�t�g�Ɍ��炸�A���܂��܂ȃf�[�^�x�[�X�\�t�g�ł��ǂݍ��ނ��Ƃ��\�ł��B<br>
                        <br>
                    </dl>
				   
				   </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">���D�{��ʂŏo�͂����CSV�t�@�C���̓��e</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt>��ʂɕ\������Ă��邷�ׂẴR���e�i�ɂ��Ď��̍��ڂ��o�͂��܂��B<br>
                      <dd> 
                        <table border="1" cellspacing="1" cellpadding="5" width=500>
                          <tr> 
                            <td bgcolor="#FFFFFF">�D��, VoyageNo, �׎�, �D��, BL No., 
                              �R���e�iNo., �w�藤�^�Ǝ�, �d�o�`���݊�������, �O�`���݊�������, CY���݌v��, CY���ݗ\�莞��, 
                              CY���݊�������, CY������������, CY���o��������, SY�\�񎞍�, SY���o��������, 
                              �q�ɓ����w������, �q�ɓ�����������, �f�o����������, ��R���ԋp����, �����m�F�\�莞��, 
                              �����m�F��������, ���A�����u, �ʔ���, �ʊ�/�ېŗA��, DO���s, �t���[�^�C��, ���o��, 
                              �T�C�Y, ����, ���[�t�@�[, ���d��, �댯��, ���o�^�[�~�i����, �X�g�b�N���[�h���p, �ԋp��, 
                              �d�o�`, �O�` </td>
                          </tr>
                        </table>
                        <br>
                      <dt>���CSV�t�@�C���̗�̂悤�ɂP�s�ڂ����ږ��łQ�s�ڈȍ~���l�ƂȂ�܂��B<BR>
                    </dl>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>���DCSV�t�@�C���o�͂̕��@</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dt> ��ʏ�́wCSV�t�@�C���o�́x�{�^�����������Ƃŕۑ���ƕۑ��t�@�C�������w�肷���ʂ��\������܂��B<br>
                    <dd> 
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td valign="top" nowrap><font color="#FF0033">�y��z</font></td>
                          <td> 
                                  <form>
                                    <input type=button value=" CSV�t�@�C���o��" name="�{�^��">
                                  </form>
                            
                          </td>
                        </tr>
                      </table><br>
                    <dt>�ۑ���ƕۑ��t�@�C�����͂Ƃ��Ɏ��R�ł����A�t�@�C�����̊g���q�͒ʏ�A�u.csv�v�Ƃ��܂��B
                    <dd><font color="#FF0033">�y��z</font>C:\MyDocument���� abcdef.csv  �Ƃ����t�@�C�����ŕۑ����܂��B<br>
                      <br>
                  </td>
                </tr>
              </table>



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
          <td bgcolor="000099" height="10"><img src="gif/1.gif"></td>
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