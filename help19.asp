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
          <td rowspan=2><img src="gif/helpt.gif" width="506" height="73"></td>
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
                  <td nowrap><b>�C�ݓ��́|�A�o�ݕ�������</b></td>
                  <td><img src="gif/hr.gif" width="400" height="3"></td>
                </tr>
              </table>
<center>
              <table border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>���DCSV�t�@�C���]���Ƃ́H</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">���͂�������񂪑����ꍇ�A���x�����͂���͖̂ʓ|�ł��B<br>
                    �����ŁA�{�V�X�e���ł͏��𗅗񂵂��t�@�C�������A���̃t�@�C����]�����邱�Ƃł܂Ƃ߂ē��͂���@�\��p�ӂ��Ă��܂��B<br>
                    �{�V�X�e���ɓ]���ł���t�@�C���̌`���́uCSV�t�@�C���v�Ƃ������ʓI�Ȃ��̂ł��B<br>
                    ���́uCSV�t�@�C���v���쐬���]�����s���菇���ȉ��ɐ������܂��B<br>
                    &nbsp; </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>���D�K�v�ȃA�v���P�[�V����</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt>CSV�t�@�C���̍쐬��Windows�t���̃������ŉ\�ł��B���邢�́AEXCEL�ō쐬����CSV�t�@�C���`���ŕۑ����邱�Ƃ��\�ł��B<br>
                    </dl>
				   
				   </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">���DCSV�t�@�C���̍쐬</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                    <td width="575"> 
                      <dl> 
                        <dt>�O�q�̃A�v���P�[�V�������g���āA�R�[���T�C���AVoyage No.�A�׎�R�[�h�E�E�E�̏��ɂЂƂЂƂ̒l���J���}�u,�v�ŋ�؂�Ȃ���P�s�ɂP�Z�b�g�̏����L�q���܂��B<br>
                          <table border="1" cellspacing="1" cellpadding="5" width=500>
                            <tr> 
                              <td bgcolor="#FFFFFF" nowrap><font size="2">A1284, 
                                B3567, 22345, 123567890, book1345, ehk, 2002/3/12/14/5, 
                                2002/3/12, 40, SD, 96, ����VP, KCVBY<br>
                                F8976, D7909, 88293, 334455666, book3746, yeg, 
                                2002/3/18/13/45, 2002/3/18, 20, RF, 86, ����VP, 
                                GFDSH</font></td>
                            </tr>
                          </table>
                          <br>
                        <dt>1�s���̍��ڂ̏ڍ׎d�l<BR>
                        <dd> 
                          <table width="100" border="1" cellspacing="0" cellpadding="2" bgcolor="#FFFFFF">
                            <tr bgcolor="#99aaFF" align="center"> 
                              <td nowrap><b><font color="#333333">����</font></b></td>
                              <td nowrap><b><font color="#333333">��</font></b></td>
                              <td nowrap><b><font color="#333333">���͎d�l</font></b></td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�D���i�R�[���T�C���j</td>
                              <td nowrap>A1284</td>
                              <td nowrap>���p�啶���p����7���ȓ�</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>Voyage No.</td>
                              <td nowrap>B3567</td>
                              <td nowrap>���p�啶���p����12���ȓ�<br>
                                ����L���܂ށi'-'�A'/'�Ȃǁj</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�׎�R�[�h</td>
                              <td nowrap>22345</td>
                              <td nowrap>JUSTPRO�̃R�[�h(���p����5��)</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�׎�Ǘ��ԍ�</td>
                              <td nowrap>123567890</td>
                              <td nowrap>���p�p����10���ȓ�</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>Booking No.</td>
                              <td nowrap>book1345</td>
                              <td nowrap>���p�p����20���ȓ�</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�w�藤�^�Ǝ҃R�[�h</td>
                              <td nowrap>ehk</td>
                              <td nowrap>���p�p��3��</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>��R���q�ɓ����w�����<br>
                                �i�N���������j</td>
                              <td nowrap>2002/3/12/14/5</td>
                              <td nowrap>�E�N�F����4��<br>
                                �E���̑��F����2���i'01'��'1'�̗����̕\���ɑΉ��j<br>
                                �E�ȏ�𔼊p�X���b�V���u/�v�ŋ�؂�B<br>
                                �E�l�������ꍇ�̓X���b�V���������c���i�u//�v�j</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>CY�����w���</td>
                              <td nowrap>2002/3/12</td>
                              <td nowrap>�i����j</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�T�C�Y</td>
                              <td nowrap>40</td>
                              <td nowrap>����2��</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�^�C�v</td>
                              <td nowrap>SD</td>
                              <td nowrap>�p��2��</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�n�C�g</td>
                              <td nowrap>96</td>
                              <td nowrap>����2��</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>��R���s�b�N�ꏊ</td>
                              <td nowrap>����VP</td>
                              <td nowrap>20byte�i�S�p�Ȃ�P�O�����A���p�Ȃ�Q�O�����ȓ��j</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>�q�ɗ��́i��R���͂���j</td>
                              <td nowrap>KCVBY</td>
                              <td nowrap>5byte�i�ʏ피�p�p�����T�����ȓ��ŁB�S�p�Ȃ�Q�����܂Łj</td>
                            </tr>
                          </table>
                        <dd>(*)�S���ڋ��ʁE�E�E���p�J�i�͋֎~
                          <p> 
                        <dt>�t�@�C�����͉��ł����܂��܂��񂪁A�g���q�͒ʏ�u.csv�v�Ƃ��܂��B�ۑ�������R�ł� 
                        <dd><font color="#FF0033">�y��z</font>C:\MyDocument���� abcdef.csv 
                          �Ƃ����t�@�C�����ŕۑ����܂��B 
                      </dl>
                  </td>
                </tr>
                <tr> 
                    <td colspan="2" bgcolor="#99ccFF"><b>���DCSV�t�@�C���̓]��</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                    <td width="575"> 
                      <ul>
                        <li>��ʏ��CSV�t�@�C���]�����N���b�N����Ǝ��̂悤��CSV�t�@�C�����w�肷���ʂ��\������܂��B<br>
                          <table border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td valign="top" nowrap><font color="#FF0033">�y��z</font></td>
                              <td> 
                                <table border="1" cellspacing="1" cellpadding="5">
                                  <tr> 
                                    <td bgcolor="#FFFFFF" align="center"> 
                                      <form>
                                        <table border="1" cellspacing="0" cellpadding="2">
                                          <tr> 
                                            <td bgcolor="#000099" nowrap> <font color="#FFFFFF"><b>CSV�t�@�C����</b></font> 
                                            </td>
                                            <td nowrap> 
                                              <input name=csvfile size=30 accept="text/css">
                                            </td>
                                            <td nowrap> 
                                              <input type=button value="�Q��..." name="�{�^��">
                                            </td>
                                          </tr>
                                        </table>
                                        <input type=button value=" ��  �M " name="�{�^��">
                                      </form>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table><br>
                        <li>�󗓂ɍ쐬����CSV�t�@�C���̃t���p�X���L�q���܂��B <br>
                          <font color="#FF0033">�y��z</font>�쐬��̏ꍇ�́uC:\MyDocument\abcdef.csv�v�ƋL�q���܂��B<br>
                        <li>����͂���̂��ʓ|�ȏꍇ�́A�m�Q��...�n�{�^���������ƃt�@�C����I�������ʂ��o�܂��̂ŁA�ۑ���̃t�H���_�ƃt�@�C�������ɑI�����Ă������ƂŃt�@�C�����������I�ɓ��͂���܂��B<br>
                        <li>�Ō�Ɂm���M�n�{�^���������܂��B<br>
                        <li>���͌��ʂ͒ʏ�̉�ʂŕ\������܂��B 
                          <p> 
                          <table border="1" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td bgcolor="#FF9933" nowrap valign="top">����</td>
                              <td bgcolor="#FFFFFF">�t�@�C���̍쐬���K���ǂ���ł��Ă��Ȃ��ƃV�X�e���͓��e��ǂݏo�����Ƃ��ł����G���[��\�����܂��B���̏ꍇ�́A�t�@�C���̓��e���������A�C��������ēx���M���s���Ă��������B</td>
                            </tr>
                          </table>
                      </ul>
                    </td>
                </tr>
              </table>

</center>

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