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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------����������--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="../gif/helpt.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
  <tr>
	<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
' End of Addition by seiko-denki 2003.07.07
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.07
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%> &gt; �w���v
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.07 -->
		<BR>
		<BR>
		<BR>
        <table>
          <tr>
            <td align="center"> 
              <table>
                <tr> 
                  <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                  <td nowrap> <b><font color="#000000">�A���R���e�i���Ɖ�L�[����</font></b>&nbsp;&nbsp;</td>
                  <td><img src="../gif/hr.gif"></td>
                </tr>
              </table>

              <table border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">���DCSV�t�@�C���]���Ƃ́H</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">�Q�Ƃ������R���e�iNo.��BL No�D�������ꍇ�A���x�����͂��Č��������s����͖̂ʓ|�ł��B<br>
                    �����ŁA�{�V�X�e���ł͎Q�Ƃ������R���e�iNo.�A�܂��́ABL No�D�𗅗񂵂��t�@�C�������A���̃t�@�C����]�����Ă܂Ƃ߂Č������s���@�\��p�ӂ��Ă��܂��B<br>
                    �{�V�X�e���ɓ]���ł���t�@�C���̌`���́uCSV�t�@�C���v�Ƃ������ʓI�Ȃ��̂ł��B<br>
                    ���́uCSV�t�@�C���v���쐬���]�����s���菇���ȉ��ɐ������܂��B<br>
                    &nbsp; </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">���D�K�v�ȃA�v���P�[�V����</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">CSV�t�@�C���̍쐬��Windows�t���̃������ŉ\�ł��B���邢�́AEXCEL�ō쐬����CSV�t�@�C���`���ŕۑ����邱�Ƃ��\�ł��B<br>
                    &nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">���DCSV�t�@�C���̍쐬 
                    </font><font color="#666666"> </font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt><b>�i�P�j�����̃R���e�iNo.�ŎQ�Ƃ������ꍇ</b> 
                      <dd>�O�q�̃A�v���P�[�V�������g���ĂP�s�ɂP�̃R���e�iNo.���L�q���A�ړI�̃R���e�iNo.�̐������s�����܂��B<br>
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td valign="top" nowrap><font color="#FF0033">�y��z</font></td>
                            <td> 
                              <table border="1" cellspacing="1" cellpadding="5" width=300>
                                <tr> 
                                  <td bgcolor="#FFFFFF">KYGU2234455<BR>
                                    GFDU2556379<BR>
                                    FGYU9882567<br>
                                    <br>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                      <dd>�t�@�C�����͉��ł����܂��܂��񂪁A�g���q�͒ʏ�u.csv�v�Ƃ��܂��B�ۑ�������R�ł��B<br>
                        <font color="#FF0033">�y��z</font>C:\MyDocument���� abcdef.csv  �Ƃ����t�@�C�����ŕۑ����܂��B<br>
                        <br>
                      <dt><b>�i�Q�j������BL No�D�ŎQ�Ƃ������ꍇ</b><br>
                      <dd>�R���e�iNo.�̏ꍇ�Ɠ��l�ɂP�s�ɂP��BL No�D���L�q���A�ړI��BL No�D�̐������s�����܂��B 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td valign="top" nowrap><font color="#FF0033">�y��z</font></td>
                            <td> 
                              <table border="1" cellspacing="1" cellpadding="5" width=300>
                                <tr> 
                                  <td bgcolor="#FFFFFF">BL12546<BR>
                                    BL88976<br>
                                    <br>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                        <BR>
                        �t�@�C�����̋K�������l�ł��B<BR>
                        <table border="1" cellspacing="0" cellpadding="3">
                          <tr> 
                            <td bgcolor="#FF9933">����</td>
                            <td bgcolor="#FFFFFF">�P��CSV�t�@�C���̒��ɃR���e�iNo.��BL No�D�����݂����邱�Ƃ͂ł��܂���B 
                            </td>
                          </tr>
                        </table>
                        <br>
                    </dl>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">���DCSV�t�@�C���̓]��</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dt> ��ʏ��CSV�t�@�C���]�����N���b�N����Ǝ��̂悤��CSV�t�@�C�����w�肷���ʂ��\������܂��B<br>
                    <dd> 
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
                      </table>
                    <dt> 
                      <ul>
                        <li>�󗓂ɍ쐬����CSV�t�@�C���̃t���p�X���L�q���܂��B<br>
                          <font color="#FF0033">�y��z</font>�쐬��̏ꍇ�́uC:\MyDocument\abcdef.csv�v�ƋL�q���܂��B<br>
                        <li>����͂���̂��ʓ|�ȏꍇ�́A�m�Q��...�n�{�^���������ƃt�@�C����I�������ʂ��o�܂��̂ŁA�ۑ���̃t�H���_�ƃt�@�C�������ɑI�����Ă������ƂŃt�@�C�����������I�ɓ��͂���܂��B<br>
                        <li>�Ō�Ɂm���M�n�{�^���������܂��B<br>
                        <li>�������ʂ͒ʏ�̉�ʂŕ\������܂��B
                          <p> 
                          <table border="1" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td bgcolor="#FF9933" nowrap valign="top">����</td>
                              <td bgcolor="#FFFFFF">�t�@�C���̍쐬���K���ǂ���ł��Ă��Ȃ��ƃV�X�e���͓��e��ǂݏo�����Ƃ��ł����G���[��\�����܂��B���̏ꍇ�́A�t�@�C���̓��e���������A�C��������ēx���M���s���Ă��������B</td>
                            </tr>
                          </table>
                      </ul>
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
	      <td valign="bottom" align="right"><a href="index.html"><img src="../gif/b-home.gif" border="0" width="270" height="23" usemap="#map"></a></td>
        </tr>
        <tr>
          <td bgcolor="000099" height="10"><img src="../gif/1.gif"></td>
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