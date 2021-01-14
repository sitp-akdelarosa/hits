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
<!-------------ここから画面--------------------------->
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
					<option value='#'>コンテナ情報照会 </option>
					<option value='../userchk.asp?link=expentry.asp'>├ 輸出コンテナ情報照会 </option>
					<option value='../userchk.asp?link=impentry.asp'>└ 輸入コンテナ情報照会 </option>
					<option value='#'>各社入力画面</option>
					<option value='../userchk.asp?link=nyuryoku-in1.asp'>├ 船社/ターミナル入力 </option>
					<option value='../userchk.asp?link=nyuryoku-kaika.asp'>├ 海貨入力 </option>
					<option value='../userchk.asp?link=nyuryoku-te.asp'>├ ターミナル入力 </option>
					<option value='../userchk.asp?link=rikuun1.asp'>└ 陸運入力</option>
					<option value='../userchk.asp?link=sokuji.asp'> 即時搬出システム </option>
					<option value='../userchk.asp?link=hits.asp'>ストックヤード利用システム</option>
					<option value='../userchk.asp?link=terminal.asp'>ゲート前映像・混雑状況照会 </option>
					<option value='../userchk.asp?link=request.asp'>利用者アンケート・Ｑ＆Ａ</option>
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
              Top &gt; 実験の結果</font> </td>
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
                  <td nowrap><b><font color="#000000">海陸一貫物流情報システムの目的</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;近年増加が著しい輸出入コンテナ輸送の一層の効率化を図るため、海陸一貫物流情報システムについて検討し、実証実験を行いました。<br>システムの具体的な目的は次のとおりです。
<table border="0">
<tr>
	<td align="left" align="center"><b>（１）</b></td>
	<td align="left" valign="top"><b>貨物の位置情報及び通関等の手続情報の共有による業務の効率化</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">（例：事務の効率化、荷主の生産工程、販売過程の最適化、トラック運行の効率化等）</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>（２）</b></td>
	<td align="left" valign="top"><b>コンテナ輸送の時間短縮</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">（例：即時搬出システムにより、船から降ろされたコンテナを即時に搬出する）</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>（３）</b></td>
	<td align="left" valign="top"><b>コンテナターミナル周辺道路の渋滞解消</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">（例：カメラ映像による道路の混雑状況やターミナル内所要時間を確認してトラックを配車する）</td>
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
                    <td nowrap><b><font color="#000000">海陸一貫物流情報システムの構成</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;開発したシステムは、次の項目から構成されています。
<table border="0">
<tr>
	<td align="left" align="center"><b>（１）</b></td>
	<td align="left" valign="top"><b>輸出入コンテナ情報照会</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;コンテナ番号、ブッキング番号、B/L番号によって輸出入コンテナの情報を照会します。これにより、コンテナの位置、手続きの状況が確認できます。</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr><p><td colspan="2" align="center"><b>海陸一貫物流情報システムのイメージ図</b><img src="./sys_img.gif"></p></td></tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>（２）</b></td>
	<td align="left" valign="top"><b>作業情報システム</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;会社毎に定めた会社コードを利用することにより、多数のコンテナを扱う物流関係者が自社に関係する全てのコンテナを船別、関係会社別等に分類して照会できるとともに、関係する会社との指示、確認等の作業の情報伝達を行うことができます。<br>
&nbsp;また、即時搬出システム<sup><small>＊</small></sup>、空コンピックアップシステム<sup><small>＊＊</small></sup>も組み込まれています。
<p><small><sup>＊</sup>即時搬出システム：特に急ぐ輸入コンテナ貨物について、コンテナを船から降ろしたらすぐにターミナルから運び出すための手続きを行うシステム。<br>
<sup>＊＊</sup>空コンピックアップシステム：輸出貨物を詰めるための空コンテナをピックアップするとともに、倉庫へ運び、コンテナヤードへ搬入するよう指示し、確認する作業を関係者間で円滑に行うためのシステム。</small>
</p></td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>（３）</b></td>
	<td align="left" valign="top"><b>その他（ゲート前映像、ターミナル混雑状況照会）</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;ターミナルゲート前のカメラ映像、ターミナル内所要時間（ゲート入場～出場）情報を照会できます。これにより、ターミナルの混雑状況が確認できます。<br>この開発したシステムは、インターネットによりパソコンで利用できるとともに、携帯電話でもコンテナ搬出許可、ゲート前映像及びターミナル内所要時間を照会できます。</td>
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
                    <td nowrap><b><font color="#000000">実証実験の結果</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;博多港を具体例として海陸一貫物流情報システムの検討を行うとともに、博多港において平成14年2月18日から3月15日まで実証実験を行いました。<br>&nbsp;実証実験の結果は次のとおりです。
<table border="0">
<tr>
	<td align="left" align="center"><b>（１）</b></td>
	<td align="left" valign="top"><b>システムの利用状況</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;実験期間中のシステム利用状況は以下のとおりで、ゲート前映像、ターミナル内所要時間照会や輸出入コンテナ情報照会を中心に活用され、有効であることが確認できました。</td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">
		<table border="0">
		<tr>
			<td align="left" align="center">○</td>
			<td align="left" valign="top">パソコンを利用したアクセス件数</td>
		</tr>
		<tr>
			<td><br></td>
			<td align="left" valign="top">実証実験期間中の合計アクセス数　　18,897件<br>
（参考：期間中にターミナルへの搬入された実入りコンテナ数は輸出6,653本、輸入8,997本、計15,650本）<br>
平日の平均アクセス数　　約1,000件
			<table border="0">
				<tr>
					<td align="left" valign="top">実証実験期間中の合計アクセス数</td>
					<td width="20"><br></td>
					<td align="left" valign="top">418,897件</td>
				</tr>
				<tr>
					<td align="left" valign="top" colspan="3">（参考：期間中にターミナルへの搬入された実入りコンテナ数は輸出6,653本、輸入8,997本、計15,650本）</td>
				</tr>
				<tr>
					<td align="left" valign="top">平日の平均アクセス数</td>
					<td width="20"><br></td>
					<td align="left" valign="top">約1,000件</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td align="left" align="center">○</td>
			<td align="left" valign="top">携帯電話を利用したアクセス件数</td>
		</tr>
		<tr>
			<td><br></td>
			<td align="left" valign="top">
			<table border="0">
				<tr>
				<td align="left" valign="top">実証実験期間中の合計アクセス数<br>平日の平均アクセス数</td>
				<td width="20"><br></td>
				<td align="left" valign="top">4,886件<br>約260件</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td align="left" align="center">○</td>
			<td align="left" valign="top">利用頻度の高い項目</td>
		</tr>
		<tr>
			<td><br></td>
			<td align="left" valign="top">
			（パソコン利用の場合）
			<table border="0">
				<tr>
				<td align="left" valign="top">・ゲート前映像<br>・ターミナル内所要時間<br>・輸入コンテナ情報照会<br>・輸出コンテナ情報照会</td>
				<td width="20"><br></td>
				<td align="right" valign="top">3,146件<br>1,835件<br>2,142件<br>947件</td>
				</tr>
			</table>
			（携帯電話利用の場合）
			<table border="0">
				<tr>
				<td align="left" valign="top">・コンテナ搬出許可照会<br>・ゲート前映像</td>
				<td width="20"><br></td>
				<td align="right" valign="top">1,329件<br>537件</td>
				</tr>
			</table>
		</tr>
	</table>
	</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>（２）</b></td>
	<td align="left" valign="top"><b>実証実験の効果</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;本システムの実証実験は、国が主体となって輸出入コンテナ輸送に関する通関等の手続き情報を含めた情報の共有化を目指す初めての試みであり、実験によって以下のような効果が確認されました。</td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">
		<table border="0">
		<tr>
			<td align="left" align="center" valign="top">○</td>
			<td align="left" valign="top">貨物の位置情報及び通関等の手続情報の共有による業務の効率化</td>
		</tr>
		<tr>
			<td align="left" valign="top" colspan="2">
			<UL>
			<LI>輸出入コンテナに関する位置、手続き情報を本システムで一元的に照会可能となり、輸出入関係者の業務効率化に有効であることが確認されました。特にトラックの運行効率化には有効でした。
			<LI>本システムの活用により、関係者が情報を電子的に取得することが可能となりました。得られた情報の積極的な活用によるワンインプット化、ペーパーレス化が促進されるものと期待されます。
			</UL>
			</td>
		</tr>
		<tr>
			<td align="left" align="center" valign="top">○</td>
			<td align="left" valign="top">コンテナ輸送の時間短縮</td>
		</tr>
		<tr>
			<td align="left" valign="top" colspan="2">
			<UL>
			<LI>即時搬出システムにより、事前に所定の通関等の手続き条件を満たした貨物をターミナル到着後速やかに搬出することが可能となりました。なお、即時搬出システムについては対象コンテナが少なかったため十分なデータが得られておりません。
			</UL>
			</td>
		</tr>
		<tr>
			<td align="left" align="center" valign="top">○</td>
			<td align="left" valign="top">コンテナターミナル周辺道路の渋滞解消</td>
		</tr>
		<tr>
			<td align="left" valign="top" colspan="2">
			<UL>
			<LI>博多港では既に平成12年11月から輸入貨物のターミナル搬出可否情報を携帯電話等により照会してターミナル周辺の渋滞解消に大きな効果をあげていましたが、本実証実験ではより詳細な情報を加えるとともに、ターミナルゲート前カメラ映像とターミナル内所要時間を照会可能とし、トラックの効率的な配車やターミナル混雑の確認等に有効でした。
			</UL>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td height="10" colspan="2"><br></td></tr>
<tr>
	<td align="left" align="center"><b>（３）</b></td>
	<td align="left" valign="top"><b>今後の課題</b></td>
</tr>
<tr>
	<td><br></td>
	<td align="left" valign="top">&nbsp;今回実験を行った海陸一貫物流情報システムが今後より使いやすいシステムとなるように、利用者から出された次のような要望や課題について改良する必要があります。</td>
</tr>
<tr>
	<td align="left" valign="top" colspan="2">
	<UL>
		<LI>コンテナ情報照会について、メニューや画面を利用しやすい構成にする。
		<LI>作業情報システムについて、システムの利用方法や運用ルールの徹底を図る。
		<LI>即時搬出システムについては範囲を広げて、包括保税輸送許可貨物等も対象にする。
		<LI>空コンピックアップシステムについて、輸出コンテナの作業情報システムと一体化させ、関係者への指示、確認といった業務の流れに合わせて利用しやすくする。
		<LI>携帯電話での照会や入力等については、今後の機器や通信サービスの発展に対応してより使いやすいものとする。
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
                    <td nowrap><b><font color="#000000">アクセス集計表</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="360" height="3"></td>
                </tr>
              </table>

			  <center>
				<table border=0 cellpadding=1 cellspacing=1 width=80%>
				  <tr>
					<td align=left>
&nbsp;<a href="logview.asp">クリックすると、日付ごとのアクセス集計表を見ることができます。</a>
					</td>
				  </tr>
				  <tr>
					<td align=left>
&nbsp;<a href="logija.asp">クリックすると、日付ごとのアクセス集計表（携帯）を見ることができます。</a>
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
<!-------------画面終わり--------------------------->
<map name="map"> 
  <area shape="poly" coords="20,0,152,0,134,22,0,22" href="JavaScript:FancBack()">
  <area shape="poly" coords="154,0,136,22,284,22,284,0" href="http://www.hits-h.com/index.asp">
</map>
</body>
</html>