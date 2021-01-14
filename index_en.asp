<% @LANGUAGE = VBScript %>
<!--#include File="./inform/common.inc"-->
<%
	Dim param(2), fso, fod, fic, home_path, rel_path
	Dim cnt, file_info, i, j, k, work
	Dim file_num,fil
	'''iniファイルの値の読み込み
	getIni param

	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set fod=fso.GetFolder(param(0))
	set fic=fod.Files
	cnt=0
	For Each fil In fic
		cnt=cnt+1
	Next

	ReDim file_info(cnt)
	'''ホームディレクトリの絶対物理パスの取り出し
	home_path=Request.ServerVariables("APPL_PHYSICAL_PATH")
	'''ホームディレクトリの相対パス
	'rel_path="/" & Replace(Right(param(0),len(param(0))-len(home_path)),"\","/")
	'Right(param(0).ToString(),len(param(0).ToString())) - len(home_path)
	rel_path="/" & Replace(home_path,"\","/")

	i=0
	'''file_info配列にファイルの作成日と名前を格納
	For Each fil In fic
		if(DateDiff("d",fil.DateLastModified,Date)<=CInt(param(1))) then '''今日−作成日<=特定期間
			file_info(i)=Left(fil.DateLastModified,4) & "年" & Mid(fil.DateLastModified,6,2) & "月" & Mid(fil.DateLastModified,9,2) & "日" & "|" & Mid(fil.DateLastModified,12,8) & "|" & fil.Name & "|1"
		else		'''今日−作成日>特定期間
			file_info(i)=Left(fil.DateLastModified,4) & "年" & Mid(fil.DateLastModified,6,2) & "月" & Mid(fil.DateLastModified,9,2) & "日" & "|" & Mid(fil.DateLastModified,12,8) & "|" & fil.Name & "|0"
		end if
		i=i+1
	Next
	file_num=i
	f=Array(0,0)
	ReDim f(file_num,3)
	'''作成日の新しいものがより上に表示されるようにソートする
	For i = 0 To UBound(file_info) - 1
		For j = i + 1 To UBound(file_info)
			If StrComp(file_info(i),file_info(j),1)<0 Then '''file_info(i)がfile_info(j)より小さい
				work=file_info(i)
				file_info(i)=file_info(j)
				file_info(j)=work
			End If
		Next
	Next
%>	

<!DOCTYPE html>
<html lang="ja">
<head>
	<meta http-equiv="Content-Type" content="text/html ; charset=Shift_JIS" />
	<meta http-equiv="Content-Style-Type" content="text/css" />
	<meta http-equiv="Content-Script-Type" content="text/javascript" />
	<title>HAKATA PORT LOGISTICS  IT SYSTEM</title>
	<meta name="description" content="説明文を入れて下さい" />
	<meta name="keywords" content="キーワードを入れて下さい" />
	<meta name="author" content="著作者を入れて下さい" />

	<% '2013/09/27 Mw.Tanaka Upd-S CSSキャッシュ対策 %>
	<!--  <link rel="stylesheet" href="newsite/css/main.css" type="text/css" /> -->
	<link rel="stylesheet" href="newsite/css/main.css?ver=150423" type="text/css" />
	<% '2013/09/27 Mw.Tanaka Upd-E CSSキャッシュ対策 %>

	<!-- // Edited by AK.DELAROSA 2021/01/11 Start -->
	<script src="newsite/scripts/createjs.min.js" type="text/javascript"></script>
	<script src="newsite/scripts/sc_1.js" type="text/javascript"></script>
	<script src="newsite/scripts/top.js" type="text/javascript"></script>
	<script src="scripts/pages/index.js" type="text/javascript"></script>
	<script src="newsite/scripts/flash.js" type="text/javascript"></script>
	<!-- // Edited by AK.DELAROSA 2021/01/11 End -->
</head>

<body onload="init(); move_icon();"> <!-- init function is from this document / move_icon is from newsite/scripts/sc_1.js -->


	<div id="main-container">
		<div id="header" style="background-image: url(newsite/img/header_en.jpg);background-repeat: no-repeat;">
			<ul id="login">
				<!-- Y.TAKAKUWA Upd-S 2015-01-27 -->
				<!--<a href="english/userchk.asp?link=index_en.asp" target="_top"><img src="newsite/img/login_btn_en.gif" border="0"></a>-->
				<li style="display:inline;margin-right:5px;">
					<a href="index.asp" target="_top"><img src="images/日本語.bmp"  height="26" width="80"  border="0"></a>
				</li>
				<li style="display:inline;margin-right:30px;">
					<a href="index_ch.asp" target="_top"><img src="chinese/image/中文簡体.bmp"  height="26" width="80"  border="0"></a>
				</li>
				<li style="display:inline">
					<a href="english/userchk.asp?link=index_en.asp" target="_top"><img src="newsite/img/login_btn_en.gif" border="0"></a>
				</li>
				<!-- Y.TAKAKUWA Upd-E 2015-01-27 -->
			</ul>

			<ul id="top_navi">
				<img src="images/headmenu_en.png" border="0" usemap="#headmenu_en">
				<map name="headmenu_en">
					<!-- Y.TAKAKUWA Upd-S 2015-03-14 -->
					<!--
					<area href="index.asp" target="_top" shape="rect"  coords="0,0,48,13">
					<area href="http://www.hits-h.com/English/info.html" target="_top" shape="rect"  coords="55,0,141,13">
					-->
					<!-- Y.TAKAKUWA Upd-S 2015-04-27 -->
					<area href="english/userchk.asp?link=touroku/index.html" target="new_window" shape="rect"  coords="55,0,214,13">
					<area href="JavaScript:openwin()" shape="rect"  coords="221,0,316,13">
					<!--<area href="index.asp" target="_top" shape="rect"  coords="267,0,310,13">-->
					<area href="http://www.hits-h.com/English/info.html" target="_top" shape="rect"  coords="325,0,413,13">
					<!-- Y.TAKAKUWA Upd-E 2015-04-27 -->
					<!-- Y.TAKAKUWA Upd-E 2015-03-14 -->
				</map>
			</ul>
		</div>
   
		<div id="left_block">
			<div id="topics">
			
				<div id="marquee" align="center"> 
					<span style="text-align:center;font-weight:bold;font-size:14px;">Welcome to HiTS Ver.3</span>
					<!--     
					<iframe src="denbun.asp" height="21" width="265" scrolling="no" frameborder="0" name="denbun"></iframe>
					-->
				</div>
				
				<!-- Y.TAKAKUWA Del-S 2015-01-29 -->
				<!--
				<div id="topic_text" style="width:270px;overflow-x:auto;">
					<Table border="0" cellspacing="0" cellpadding="0" width="250">							
					<% 
						If file_num >0 then 
							For k=0 to file_num-1
								file_data=split(file_info(k),"|")
								j=0
								for each fd in file_data
									f(k,j)=fd
									j=j+1
								next
								response.write "<tr>"
								if f(k,3)=1 then
									response.write "<td width='10'>"
									response.write "<img src='../inform/images/new2.gif' border='0'>"
									response.write "</td>"					
								end if
								if f(k,3)<>1 then
									response.write "<td colspan='2' width='100%'>"
								else
									response.write "<td width='100%'>"
								end if
								
								response.write "<p>"
								response.write "<a href=" & rel_path & f(k,2) & " target='_blank'>" & f(k,0) & " " & left(f(k,2),len(f(k,2))-4) & "</a>"
								response.write "</p>"					
								response.write "</td>"
								response.write "</tr>"
							Next	
						End If
					%>		
					</Table> 
				</div>
				-->
				<!-- Y.TAKAKUWA Del-E 2015-01-29 -->
			</div>
		
			<div id="links">
				<ul id="btn">
					<li><a href="http://www.jphkt.co.jp/" target="new_window"><img src="english/images/LINK_1_en.png" alt="博多港運" /></a></li>
					<li><a href="http://www.sogo-unyu.co.jp/" target="new_window"><img src="english/images/LINK_2_en.png" alt="相互運輸" /></a></li>
					<li><a href="http://www.nittsu.co.jp/" target="new_window"><img src="english/images/LINK_3_en.png" alt="日本通運" /></a></li>
					<li><a href="http://www.geneq.co.jp/" target="new_window"><img src="english/images/LINK_4_en.png" alt="ジェネック" /></a></li>
					<li><a href="http://www.kamigumi.co.jp" target="new_window"><img src="english/images/LINK_5_en.png" alt="上組" /></a></li>
					<li><a href="http://www.mitsubishi-logistics.co.jp/" target="new_window"><img src="english/images/LINK_6_en.png" alt="三菱倉庫" /></a></li>
					<li><a href="http://www.hakatako-futo.co.jp/" target="new_window"><img src="english/images/LINK_7_en.png" alt="博多港ふ頭" /></a></li>
					<li><a href="http://port-of-hakata.city.fukuoka.lg.jp/" target="new_window"><img src="english/images/LINK_8_en.png" alt="福岡市港湾局" /></a></li>
					<li><a href="http://www.gct.com.cn/" target="new_window"><img src="newsite/img/LINK_9.png" alt="GUANGZHOU CONTAINER TERMINAL" /></a></li>
					<li><a href="http://www.gnict.com/cn/default.jsp" target="new_window"><img src="newsite/img/LINK_10.png" alt="NANSHA STEVEDORING" /></a></li>
					<li><a href="http://www.sctcn.com/" target="new_window"><img src="newsite/img/LINK_11.png" alt="SHEKOU CONTAINER TERMINALS" /></a></li>
					<li><a href="http://www.cwcct.com//cct/cct_en/publicinf/main/index.aspx" target="new_window"><img src="newsite/img/LINK_12.png" alt="CHIWAN CONTAINER TERMINAL" /></a></li>
					<li><a href="http://www.port.co.th/sitenew/en/" target="new_window"><img src="newsite/img/LINK_15.png" alt="Port Authority of Thailand" /></a></li>
					<li><a href="http://www.tpct.com.tw/" target="new_window"><img src="newsite/img/LINK_14.png" alt="TAIPEI PORT CONTAINER TERMINAL" /></a></li>
					<!--       <li><img src="newsite/img/LINK_btn.png" alt="Blank" /></li>-->
				</ul>
				<div id="linkbanner_left">
					<ul id="bnr">
						<li id="nowfas"><a href="http://www.mlit.go.jp/kowan/nowphas/" target="_blank"><img src="newsite/img/nowfas.jpg" alt="ナウファス" /></a></li>
						<li id="mitsui"><a href="http://www.mes.co.jp" target="_blank"><img src="newsite/img/mitsui_zosen.jpg" alt="三井造船" /></a>
							<img src="images/mitsui.gif">
						</li>
					</ul>
				</div>
			</div>
		</div>
	
		<div id="center_block">
			<p class="title"><img src="english/images/Explain_title_en2.png" title="" border="0"></p>
			<p class="description"><img src="english/images/Explain_en2.png" title="" border="0"></p>
			<div id="global_menu">
				<ul>
					<li id="gm_exinfo_en2"><a href="english/expentry.asp" target="_top"><span>EXPORT CONTAINER INFORMATION</span></a></li>
					<li id="gm_entry_en2"><a href="english/pdf/HiTS_事前情報入力(英訳)rev1.pdf" target="_blank"><span>PRIOR INFORMATION ENTRY</span></a></li>
					<li id="gm_iminfo_en2"><a href="english/impentry.asp" target="_top"><span>IMPORT CONTAINER INFORMATION</span></a></li>
					<li id="gm_reservation_en2"><a href="english/shuttle/shuttle-rest.html" target="_top"><span>SHUTTLE RESERVATION</span></a></li>
					<li id="gm_adinfo_en2"><a href="english/arvdepinfo.asp" target="_top"><span>VESSEL ARRIVAL AND DEPARTURE INFORMATION</span></a></li>
					<li id="gm_request_en2"><a href="english/userchk.asp?link=SendStatus/sst000F.asp" target="_top"><span>IMPORT STATUS DELIVERY REQUEST</span></a></li>
					<li id="gm_cyinfo_en2"><a href="english/terminal.asp" target="_top"><span>CY CONGESTION SITUATION INFORMATION</span></a></li>
					<li id="gm_qa_en2"><a href="english/qa/index.html" target="_top"><span>Q&A</span></a></li>
				</ul>
			</div>
		</div>

		<div id="right_block">
			<!-- <script type="text/javascript" language="javascript">
				// flash({
				// 	src : 'newsite/swf/top.swf',
				// 	w   : 259,
				// 	h   : 419
				// });
			</script>
			<noscript>
				<object type="application/x-shockwave-flash" data="newsite/swf/top.swf" width="259" height="419"><param name="movie" value="newsite/swf/top.swf" /></object>
			</noscript> -->

			<!-- // Edited by AK.DELAROSA 2021/01/11 Start -->
			<img src="newsite/scripts/top.gif" id="gif_alternative" width="259" height="419" style="position: absolute; display: none;"/>

			<div id="animation_container" style="width:259px; height:419px; display: none;">
				<canvas id="canvas" width="259" height="419" style="position: absolute; display: block;"></canvas>
				<div id="dom_overlay_container" style="pointer-events:none; overflow:hidden; width:259px; height:419px; position: absolute; left: 0px; top: 0px; display: block;">
				</div>
			</div>
			<!-- // Edited by AK.DELAROSA 2021/01/11 End -->
		</div>

		<div id="linkbanner">
			<div id="zentai"><!--ここから-->
				<a href="bannerlog.asp?longid=l101&logno=01&linkname=バナー１&linkurl=http://www.hakatako-futo.co.jp/" target="_blank"><img src="images/バナー2011.gif" title="" id="gazou1" border=0  width="130" height="60"></a><!--画像１〜３の設定-->
				<!-- 2012.4.24 Mod-S MES Aoyagi バナー更新 「広告募集中」から「船舶動静検索」-->
				<a href="bannerlog.asp?longid=l102&logno=01&linkname=バナー２&linkurl=http://www.ocean-commerce.co.jp/hakata/" target="_blank" style="color:white"><img src="images/futo-kensaku.gif" title="" id="gazou2"  border=0 width="130" height="60"></a>
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="gazou2" border=0 ></a> -->
				<!-- 2012.4.24 Mod-E MES Aoyagi バナー更新 -->	
				<!-- 2014.5.7 Mod-S MES Aoyagi -->
				<!-- <a href="http://www.tcm.co.jp" target="_blank"><img src="images/banner.gif" title="" id="gazou3"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l103&logno=01&linkname=バナー３&linkurl=http://www.unicarriers.co.jp/" target="_blank"><img src="images/UNICARRIERS.gif" title="" id="gazou3"  border=0 ></a>
				<!-- 2014.5.7 Mod-E MES Aoyagi -->
				<!--2016.7.15 Mod-S MES Aoyagi バナー更新 「広告募集中」から「西邦海運」殿-->
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="gazou4"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l104&logno=01&linkname=バナー４&linkurl=http://www.seihou.jp" target="_blank" style="color:white"><img src="images/seihou.jpg" title="" id="gazou4"  border=0 width="130" height="60" alt=""></a> 
				<!-- 2016.7.15 Mod-E MES Aoyagi バナー更新 「広告募集中」から「西邦海運」殿-->
				<a href="bannerlog.asp?longid=l105&logno=01&linkname=バナー５&linkurl=http://www.seiko-denki.co.jp/" target="_blank"><img src="images/SEIKO.JPG" title="" id="gazou5"  border=0 ></a>

				<!-- 2015.5.15 Mod-S Cosmo Nogami バナー更新 GENEQ-->		
				<!-- <a><img src="images/Blankbanner3.JPG" title="" id="gazou6" border=0 ></a> -->
				<a href="bannerlog.asp?longid=l106&logno=01&linkname=バナー６&linkurl=http://www.geneq.co.jp/" target="_blank"><img src="images/GENEQ_cm.gif" title="" id="gazou6"  border=0 ></a>
				<!-- 2015.5.15 Mod-E Cosmo Nogami バナー更新 GENEQ-->		
				<!-- 2014.5.7 Mod-S MES Aoyagi -->	
				<!-- <a href="http://www.idex.co.jp" target="_blank"><img src="images/Image00002.jpg" title="" id="gazou7"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l107&logno=01&linkname=バナー７&linkurl=http://www.idex.co.jp/" target="_blank"><img src="images/Image00002.gif" title="" id="gazou7"  border=0 ></a>
				<!-- 2016.10.14 Mod-S MES Aoyagi バナー更新 「LGX」から「実証実験中」殿-->		
				<!-- <a href="bannerlog.asp?longid=l108&logno=01&linkname=バナー８&linkurl=http://www.ditp.go.th/japan/download/article/article_20160722170749.pdf" target="_blank"><img src="images/LGX16.jpg" title="" id="gazou8"  border=0 ></a> -->
				<a><img src="images/Blankbanner4.JPG" title="" id="gazou8"  border=0 ></a> 
				<!-- 2016.10.14 Mod-E MES Aoyagi バナー更新 -->

				<a href="bannerlog.asp?longid=l101&logno=01&linkname=バナー１&linkurl=http://www.hakatako-futo.co.jp/" target="_blank"><img src="images/バナー2011.gif" title="" id="b_gazou1" border=0  width="130" height="60"></a><!--画像１〜３の設定-->
				<!-- 2012.4.24 Mod-S MES Aoyagi バナー更新 「広告募集中」から「船舶動静検索」-->
				<a href="bannerlog.asp?longid=l102&logno=01&linkname=バナー２&linkurl=http://www.ocean-commerce.co.jp/hakata/" target="_blank" style="color:white"><img src="images/futo-kensaku.gif" title="" id="b_gazou2"  border=0 width="130" height="60"></a>
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="b_gazou2" border=0 ></a> -->
				<!-- 2012.4.24 Mod-E MES Aoyagi バナー更新 -->	
				<!-- 2014.5.7 Mod-S MES Aoyagi -->
				<!-- <a href="http://www.tcm.co.jp" target="_blank"><img src="images/banner.gif" title="" id="b_gazou3"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l103&logno=01&linkname=バナー３&linkurl=http://www.unicarriers.co.jp/" target="_blank"><img src="images/UNICARRIERS.gif" title="" id="b_gazou3"  border=0 ></a>
				<!-- 2014.5.7 Mod-E MES Aoyagi -->
				<!-- 2016.7.15 Mod-S MES Aoyagi バナー更新 「広告募集中」から「西邦海運」殿-->
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="b_gazou4"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l104&logno=01&linkname=バナー４&linkurl=http://www.seihou.jp" target="_blank" style="color:white"><img src="images/seihou.jpg" title="" id="b_gazou4"  border=0 width="130" height="60" alt=""></a> 
				<!-- 2016.7.15 Mod-E MES Aoyagi バナー更新 「広告募集中」から「西邦海運」殿-->
				<a href="bannerlog.asp?longid=l105&logno=01&linkname=バナー５&linkurl=http://www.seiko-denki.co.jp/" target="_blank"><img src="images/SEIKO.JPG" title="" id="b_gazou5"  border=0 ></a>
				<!-- 2015.5.15 Mod-S Cosmo Nogami バナー更新 GENEQ-->		
				<!-- <a><img src="images/Blankbanner3.JPG" title="" id="b_gazou6" border=0 ></a> -->
				<a href="bannerlog.asp?longid=l106&logno=01&linkname=バナー６&linkurl=http://www.geneq.co.jp/" target="_blank"><img src="images/GENEQ_cm.gif" title="" id="b_gazou6"  border=0 ></a>
				<!-- 2015.5.15 Mod-E Cosmo Nogami バナー更新 GENEQ-->		
				<!-- 2014.5.7 Mod-S MES Aoyagi -->	
				<!-- <a href="http://www.idex.co.jp" target="_blank"><img src="images/Image00002.jpg" title="" id="b_gazou7"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l107&logno=01&linkname=バナー７&linkurl=http://www.idex.co.jp/" target="_blank"><img src="images/Image00002.gif" title="" id="b_gazou7"  border=0 ></a>
				<!-- 2014.5.7 Mod-E MES Aoyagi -->
				<!-- 2016.10.14 Mod-S MES Aoyagi バナー更新 「LGX」から「実証実験中」殿-->		
				<!-- <a href="bannerlog.asp?longid=l108&logno=01&linkname=バナー８&linkurl=http://www.ditp.go.th/japan/download/article/article_20160722170749.pdf" target="_blank"><img src="images/LGX16.jpg" title="" id="b_gazou8"  border=0 ></a> -->
				<a><img src="images/Blankbanner4.JPG" title="" id="b_gazou8"  border=0 ></a> 
				<!-- 2016.10.14 Mod-E MES Aoyagi バナー更新 -->
			</div><!--ここまで-->
		</div>
		
		<div id="footer_en2">
			<% '2013/09/27 Mw.Tanaka Add-S %>
			<img src="english/images/footer_en.png" border="0" usemap="#footerlink">
			<map name="footerlink">
				<area href="http://www.hits-h.com/ija/" target="_top" shape="rect"  coords="95,50,250,70">
				<area href="http://www.hits-h.com/sp/index.aspx" target="_top" shape="rect" coords="330,50,520,70">
			</map>
			<% '2013/09/27 Mw.Tanaka Add-E %>
			<div id="get_flash">
				<a href="http://get.adobe.com/jp/flashplayer/" target="_blank"><img src="newsite/img/get_flashplayer.jpg" width="66" height="16" alt="get_flashplayer" /></a>
				<p><img src="english/images/Get_flash_en2.png" title=""border=0 ></p>
			</div>
			<div id="get_adobe_reader">
				<a href="http://get.adobe.com/jp/reader/" target="_blank"><img src="newsite/img/get_adobe_reader.jpg" width="58" height="16" alt="get_adobe_reader" /></a>
				<p><img src="english/images/Get_reader_en2.png" title=""border=0 ></p>
			</div>
			<p id="copyright">Copyright(c) 2010 Hakata Port Terminal Co., Ltd. All Rights Reserved.</p>
		</div>
	</div>

</body>
</html>
