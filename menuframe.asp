<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<html>
<head>
<base target="_top">
<title>�����`����IT�V�X�e��</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="hits1.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="162" height="406" border=0 cellpadding="2" cellspacing="0" mm_noconvert="TRUE">
  <tr>
    <td height="8" bgcolor="#FFFFFF" ><img src="images/transparent.gif" width="1" height="1"></td>
  </tr>
  <!-- 2007/03/21 Upd-S Maquez��ʍ��ږ|�� -->
  <%
  		if Right(Request.ServerVariables("HTTP_REFERER"),12)="index_en.asp" then
  %>
  <tr>
  	<td class="mainmenulink"><a href="userchk.asp?link=English/expentry.asp" target="_top">Container Information�iExp�j</a></td>
  </tr>
  <tr>
	  	<td class="mainmenulink"><a href="userchk.asp?link=English/impentry.asp" target="_top">Container Information�iImp�j</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top">&nbsp;</a> </td>
  </tr>
  <tr>
    <td height="3" bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top">&nbsp;</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top">&nbsp;</a></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top">&nbsp;</a></td>
  </tr>
  <tr>
    <td height="22" class="mainmenulink"><a target="_top">&nbsp; </a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top">&nbsp;</a> </td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top"><strong>Others</strong></a></td>
  </tr>
  <tr>
    <td class="menuside"><a  href="English/info.html" >terms of service </a> </td>
  </tr>
    <tr>
    <td height="19" class="menuside"><a target="_top" >&nbsp;</a> </td>
  </tr>
  <tr>
    <td height="20" class="menuside"><a  target="_top" >&nbsp; </a></td>
  </tr>
  <% else %>
   <tr>
       <td class="mainmenulink"><a href="userchk.asp?link=expentry.asp" target="_top">�R���e�i���Ɖ�i�A�o�j</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"> <a href="userchk.asp?link=impentry.asp" target="_top">�R���e�i���Ɖ�i�A���j</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"> <a href="userchk.asp?link=arvdepinfo.asp" target="_top" >�����ݏ��Ɖ�
      </a></td>
  </tr>
  <tr>
    <td height="3" bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=Shuttle/SYWB013.asp" target="_top" >�V���g���\��i���gi�sS)</a>
    </td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=predef/dmi000F.asp" target="_top" >���O������</a></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
<!-- 2006/03/26 Del-S Fujiyama ��ʃ��C�A�E�g�ύX -->
<!--
  <tr>
    <td class="mainmenulink">�����˗�</td>
  </tr>
-->
<!-- 2006/03/26 Del-E Fujiyama ��ʃ��C�A�E�g�ύX -->
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=terminal.asp" target="_top" >�b�x���G�󋵁E�f��</a></td>
  </tr>
<!-- 2006/03/26 Del-S Fujiyama ��ʃ��C�A�E�g�ύX -->
<!--
  <tr>
    <td class="mainmenulink"><a href="menuframe1.asp" target="_self">�e�Џ�����</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=sokuji.asp" target="_top" >�������o�V�X�e��</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=pickselect.asp" target="_top" >��R���s�b�N�A�b�v�V�X�e��</a></td>
  </tr>
  <tr>
    <td height="22" class="mainmenulink"><a href="menuframe2.asp" target="_self">��Ə��V�X�e��
      </a></td>
  </tr>
-->
<!-- 2006/03/26 Del-E Fujiyama ��ʃ��C�A�E�g�ύX -->
<!-- 2009/03/17 Del-S Fujiyama
  <tr>
    <td height="22" class="mainmenulink"><a href="menuframe3.asp" target="_self">�A�N�Z�X����
      </a></td>
  </tr>
     2009/03/17 Del-E Fujiyama -->
<!-- 2006/03/28 Add-S Fujiyama ��ʃ��C�A�E�g�ύX -->
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=SendStatus/sst000F.asp" target="_top">�A���X�e�[�^�X�z�M�˗�
      </a> </td>
  </tr>
<!-- 2006/03/28 Add-E Fujiyama ��ʃ��C�A�E�g�ύX -->
  <tr>
    <td bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top"><strong>���̑�</strong></a></td>
  </tr>
<!-- 2006/03/28 Del-S Fujiyama ��ʃ��C�A�E�g�ύX -->
<!--
  <tr>
    <td class="menuside"><a href="userchk.asp?link=SendStatus/sst000F.asp" target="_top">�A���X�e�[�^�X�z�M�˗�
      </a> </td>
  </tr>
-->
<!-- 2006/03/28 Del-E Fujiyama ��ʃ��C�A�E�g�ύX -->
  <tr>
    <td class="menuside"><a href="info.html">���p�K��E�Ɛӎ���
      </a> </td>
  </tr>
  <!-- 2008/10/28 Upd-S Chris -->
   <tr>
    <td height="20" class="menuside"><a href="JavaScript:openwin()">�_�E�����[�h</a></td>
  </tr> 
  <!-- 2008/10/28 Upd-E Chris -->
    <tr>
    <td height="19" class="menuside"><a href="userchk.asp?link=mainpoint.asp" target="_top" >���؎����̌���
      </a> </td>
  </tr>
  <tr>
    <td height="20" class="menuside"><a href="userchk.asp?link=touroku/index.html" target="new_window" >��ЃR�[�h�o�^�̈ē�
      </a></td>
  </tr>
  <% end if %>
<!-- 2007/03/21 Upd-EMaquez��ʍ��ږ|�� -->
<!-- 2006/03/26 Del-S Fujiyama ��ʃ��C�A�E�g�ύX -->
<!--
  <tr>
    <td height="20" class="mainmenulink"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
-->
<!-- 2006/03/26 Del-E Fujiyama ��ʃ��C�A�E�g�ύX -->
<!-- 2006/03/26 Add-S Fujiyama ��ʃ��C�A�E�g�ύX -->
  </tr>
	<td colspan="3" height="31" valign="bottom">
		<div align="center">
			<span class="header2">�e�Ђւ̃����N��</span>
		</div>
	</td>
  </tr>
  <tr>
	<td width="180" height="40" colspan="3" valign="middle" nowrap>
		<table height="37" border="0" align="center" cellpadding="1" cellspacing="1">
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="200" height="16"><font color="#000099">&#8226; <a href="http://www.jphkt.co.jp/" target="new_window"> �����`�^ (��)  </font></td>
				<td width="200" height="16"><font color="#000099">&#8226; <a href="http://www.sogo-unyu.co.jp/" target="new_window">���݉^�A (��)</a></font></td>
			</tr>
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.nittsu.co.jp/" target="new_window">���{�ʉ^ (��)</a></font></td>
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.geneq.co.jp/" target="new_window">(��) �W�F�l�b�N</a></font></td>
			</tr>
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="200" height="16"><font color="#000099">&#8226; (��) ��g </font></td>
				<td width="200" height="16"><font color="#000099">&#8226; �O�H�q�� (��)</font></td>
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.hakatako-futo.co.jp/" target="new_window">�����`�ӓ� (��)</a></font></td>
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.port-of-hakata.or.jp/" target="new_window">�����s�`�p��</a></font></td>
			</tr>
		</table>
	</td>
  </tr>

  <!-- 2007/03/21 Upd-S Marquez ��ʃ��C�A�E�g�ύX -->
  <!--
    </tr>
	<td colspan="3" height="31" valign="bottom">
		<div align="center">
			<span class="header2">�g�уA�h���X</span>
		</div>
	</td>
	<tr colspan="3" height="31" valign="bottom">
		<div align="center">
			<td width="200" height="16"> http://www.hits-h.com/ija/ </td>
		</div>
	</tr>
-->
  <tr>
	<td  height="70" colspan="3" valign="bottom" nowrap>
	<table width='100% ' height="37" border="0" align="center" cellpadding="1" cellspacing="1">
	<tr>
	 <td width='50%' align='center' valign="middle"><a href="http://www.cwcct.com//cct/cct_en/publicinf/main/index.aspx" target="new_window">
		  	<img src="images/CCT.gif"  width="80" height="60" border=></a></td>
	  <td width='50%' align='center' valign="middle"><a href="http://www.sctcn.com/english/default.aspx" target="new_window">
	  		<img src="images/SCT.jpg"  width="80" height="60" border=0></a></td>
	</tr>
	</table>
	</td>
  </tr>
  <!-- 2007/03/21 Upd-EMarquez ��ʃ��C�A�E�g�ύX -->
  <tr>
	<td colspan="3" height="70" valign="middle" align="center">
		<a href="http://www.mlit.go.jp/kowan/nowphas/"><img src="images/nowfas.gif" border="0" alt="�i�E�t�@�X"></a><img src="images/transparent.gif" width="5" height="1">
	</td>
  </tr>
<!-- 2006/03/26 Add-E Fujiyama ��ʃ��C�A�E�g�ύX -->
</table>
</body>
</html>
