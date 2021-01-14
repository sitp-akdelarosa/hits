<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<html>
<head>
<base target="_top">
<title>博多港物流ITシステム</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<link href="hits1.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="162" height="406" border=0 cellpadding="2" cellspacing="0" mm_noconvert="TRUE">
  <tr>
    <td height="8" bgcolor="#FFFFFF" ><img src="images/transparent.gif" width="1" height="1"></td>
  </tr>
  <!-- 2007/03/21 Upd-S Maquez画面項目翻訳 -->
  <%
  		if Right(Request.ServerVariables("HTTP_REFERER"),12)="index_en.asp" then
  %>
  <tr>
  	<td class="mainmenulink"><a href="userchk.asp?link=English/expentry.asp" target="_top">Container Information（Exp）</a></td>
  </tr>
  <tr>
	  	<td class="mainmenulink"><a href="userchk.asp?link=English/impentry.asp" target="_top">Container Information（Imp）</a></td>
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
       <td class="mainmenulink"><a href="userchk.asp?link=expentry.asp" target="_top">コンテナ情報照会（輸出）</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"> <a href="userchk.asp?link=impentry.asp" target="_top">コンテナ情報照会（輸入）</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"> <a href="userchk.asp?link=arvdepinfo.asp" target="_top" >着離岸情報照会
      </a></td>
  </tr>
  <tr>
    <td height="3" bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=Shuttle/SYWB013.asp" target="_top" >シャトル予約（旧ＨiＴS)</a>
    </td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=predef/dmi000F.asp" target="_top" >事前情報入力</a></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
<!-- 2006/03/26 Del-S Fujiyama 画面レイアウト変更 -->
<!--
  <tr>
    <td class="mainmenulink">検査依頼</td>
  </tr>
-->
<!-- 2006/03/26 Del-E Fujiyama 画面レイアウト変更 -->
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=terminal.asp" target="_top" >ＣＹ混雑状況・映像</a></td>
  </tr>
<!-- 2006/03/26 Del-S Fujiyama 画面レイアウト変更 -->
<!--
  <tr>
    <td class="mainmenulink"><a href="menuframe1.asp" target="_self">各社情報入力</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=sokuji.asp" target="_top" >即時搬出システム</a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=pickselect.asp" target="_top" >空コンピックアップシステム</a></td>
  </tr>
  <tr>
    <td height="22" class="mainmenulink"><a href="menuframe2.asp" target="_self">作業情報システム
      </a></td>
  </tr>
-->
<!-- 2006/03/26 Del-E Fujiyama 画面レイアウト変更 -->
<!-- 2009/03/17 Del-S Fujiyama
  <tr>
    <td height="22" class="mainmenulink"><a href="menuframe3.asp" target="_self">アクセス件数
      </a></td>
  </tr>
     2009/03/17 Del-E Fujiyama -->
<!-- 2006/03/28 Add-S Fujiyama 画面レイアウト変更 -->
  <tr>
    <td class="mainmenulink"><a href="userchk.asp?link=SendStatus/sst000F.asp" target="_top">輸入ステータス配信依頼
      </a> </td>
  </tr>
<!-- 2006/03/28 Add-E Fujiyama 画面レイアウト変更 -->
  <tr>
    <td bgcolor="#FFFFFF"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
  <tr>
    <td class="mainmenulink"><a target="_top"><strong>その他</strong></a></td>
  </tr>
<!-- 2006/03/28 Del-S Fujiyama 画面レイアウト変更 -->
<!--
  <tr>
    <td class="menuside"><a href="userchk.asp?link=SendStatus/sst000F.asp" target="_top">輸入ステータス配信依頼
      </a> </td>
  </tr>
-->
<!-- 2006/03/28 Del-E Fujiyama 画面レイアウト変更 -->
  <tr>
    <td class="menuside"><a href="info.html">利用規約・免責事項
      </a> </td>
  </tr>
  <!-- 2008/10/28 Upd-S Chris -->
   <tr>
    <td height="20" class="menuside"><a href="JavaScript:openwin()">ダウンロード</a></td>
  </tr> 
  <!-- 2008/10/28 Upd-E Chris -->
    <tr>
    <td height="19" class="menuside"><a href="userchk.asp?link=mainpoint.asp" target="_top" >実証実験の結果
      </a> </td>
  </tr>
  <tr>
    <td height="20" class="menuside"><a href="userchk.asp?link=touroku/index.html" target="new_window" >会社コード登録の案内
      </a></td>
  </tr>
  <% end if %>
<!-- 2007/03/21 Upd-EMaquez画面項目翻訳 -->
<!-- 2006/03/26 Del-S Fujiyama 画面レイアウト変更 -->
<!--
  <tr>
    <td height="20" class="mainmenulink"><a target="_top"><img src="images/transparent.gif" width="1" height="1"></a></td>
  </tr>
-->
<!-- 2006/03/26 Del-E Fujiyama 画面レイアウト変更 -->
<!-- 2006/03/26 Add-S Fujiyama 画面レイアウト変更 -->
  </tr>
	<td colspan="3" height="31" valign="bottom">
		<div align="center">
			<span class="header2">各社へのリンク先</span>
		</div>
	</td>
  </tr>
  <tr>
	<td width="180" height="40" colspan="3" valign="middle" nowrap>
		<table height="37" border="0" align="center" cellpadding="1" cellspacing="1">
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="200" height="16"><font color="#000099">&#8226; <a href="http://www.jphkt.co.jp/" target="new_window"> 博多港運 (株)  </font></td>
				<td width="200" height="16"><font color="#000099">&#8226; <a href="http://www.sogo-unyu.co.jp/" target="new_window">相互運輸 (株)</a></font></td>
			</tr>
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.nittsu.co.jp/" target="new_window">日本通運 (株)</a></font></td>
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.geneq.co.jp/" target="new_window">(株) ジェネック</a></font></td>
			</tr>
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="200" height="16"><font color="#000099">&#8226; (株) 上組 </font></td>
				<td width="200" height="16"><font color="#000099">&#8226; 三菱倉庫 (株)</font></td>
			<tr bgcolor="#99CCFF" class="menubottom">
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.hakatako-futo.co.jp/" target="new_window">博多港ふ頭 (株)</a></font></td>
				<td width="100" height="16"><font color="#000099">&#8226; <a href="http://www.port-of-hakata.or.jp/" target="new_window">福岡市港湾局</a></font></td>
			</tr>
		</table>
	</td>
  </tr>

  <!-- 2007/03/21 Upd-S Marquez 画面レイアウト変更 -->
  <!--
    </tr>
	<td colspan="3" height="31" valign="bottom">
		<div align="center">
			<span class="header2">携帯アドレス</span>
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
  <!-- 2007/03/21 Upd-EMarquez 画面レイアウト変更 -->
  <tr>
	<td colspan="3" height="70" valign="middle" align="center">
		<a href="http://www.mlit.go.jp/kowan/nowphas/"><img src="images/nowfas.gif" border="0" alt="ナウファス"></a><img src="images/transparent.gif" width="5" height="1">
	</td>
  </tr>
<!-- 2006/03/26 Add-E Fujiyama 画面レイアウト変更 -->
</table>
</body>
</html>
