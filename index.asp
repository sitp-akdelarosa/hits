<% @LANGUAGE = VBScript %>
<!--#include File="./inform/common.inc"-->
<%
	Dim param(2), fso, fod, fic, home_path, rel_path
	Dim cnt, file_info, i, j, k, work
	Dim file_num,fil
	'''ini�t�@�C���̒l�̓ǂݍ���
	getIni param

	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set fod=fso.GetFolder(param(0))
	set fic=fod.Files
	cnt=0
	For Each fil In fic
		cnt=cnt+1
	Next

	ReDim file_info(cnt)
	'''�z�[���f�B���N�g���̐�Ε����p�X�̎��o��
	home_path=Request.ServerVariables("APPL_PHYSICAL_PATH")
	'''�z�[���f�B���N�g���̑��΃p�X
	'Right(param(0).ToString(),len(param(0).ToString())) - len(home_path)
	rel_path="/" & Replace(home_path,"\","/")

	i=0
	'''file_info�z��Ƀt�@�C���̍쐬���Ɩ��O���i�[
	For Each fil In fic
		if(DateDiff("d",fil.DateLastModified,Date)<=CInt(param(1))) then '''�����|�쐬��<=�������
			file_info(i)=Left(fil.DateLastModified,4) & "�N" & Mid(fil.DateLastModified,6,2) & "��" & Mid(fil.DateLastModified,9,2) & "��" & "|" & Mid(fil.DateLastModified,12,8) & "|" & fil.Name & "|1"
		else		'''�����|�쐬��>�������
			file_info(i)=Left(fil.DateLastModified,4) & "�N" & Mid(fil.DateLastModified,6,2) & "��" & Mid(fil.DateLastModified,9,2) & "��" & "|" & Mid(fil.DateLastModified,12,8) & "|" & fil.Name & "|0"
		end if
		i=i+1
	Next
	file_num=i
	f=Array(0,0)
	ReDim f(file_num,3)
	'''�쐬���̐V�������̂�����ɕ\�������悤�Ƀ\�[�g����
	For i = 0 To UBound(file_info) - 1
		For j = i + 1 To UBound(file_info)
			If StrComp(file_info(i),file_info(j),1)<0 Then '''file_info(i)��file_info(j)��菬����
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
	<title>�����`����IT�V�X�e��</title>
	<meta name="description" content="�����������ĉ�����" />
	<meta name="keywords" content="�L�[���[�h�����ĉ�����" />
	<meta name="author" content="����҂����ĉ�����" />
	<% '2013/09/27 Mw.Tanaka Upd-S CSS�L���b�V���΍� %>
	<!-- <link rel="stylesheet" href="newsite/css/main.css" type="text/css" /> -->
	<link rel="stylesheet" href="newsite/css/main.css?ver=150423" type="text/css"/>
	<% '2013/09/27 Mw.Tanaka Upd-E CSS�L���b�V���΍� %>

	<!-- // Edited by AK.DELAROSA 2021/01/11 Start -->
	<script src="newsite/scripts/sc_1.js" type="text/javascript"></script>
	<script src="newsite/scripts/flash.js" type="text/javascript"></script>
	<script src="newsite/scripts/createjs.min.js" type="text/javascript"></script>
	<script src="newsite/scripts/top.js" type="text/javascript"></script>
	<script src="scripts/pages/index.js" type="text/javascript"></script>
	<!-- // Edited by AK.DELAROSA 2021/01/11 End -->
	
</head>


<body onload="init(); move_icon();"> <!-- init function is from this document / move_icon is from newsite/scripts/sc_1.js -->
	<div id="main-container">
		<div id="header">
			<ul id="login">
				<li></li>
				<!-- Y.TAKAKUWA Upd-S 2015-01-27 -->
				<!--<a href="userchk.asp?link=index.asp" target="_top"><img src="newsite/img/login_btn.gif" border="0"></a>-->
				<li style="display:inline;margin-right:5px;">
					<a href="index_ch.asp" target="_top"><img src="chinese/image/�����ȑ�.bmp"  height="26" width="80"  border="0"></a>
				</li>
				<li style="display:inline;margin-right:30px;">
					<a href="index_en.asp" target="_top"><img src="english/images/ENGLISH.bmp" height="26" width="80" border="0"></a>
				</li>
				<li style="display:inline">
					<a href="userchk.asp?link=index.asp" target="_top"><img src="newsite/img/login_btn.gif" border="0"></a>
				</li>
				<!-- Y.TAKAKUWA Upd-E 2015-01-27 -->
			</ul>
			<ul id="top_navi">
				<img src="images/headmenu.gif" border="0" usemap="#headmenu">
				<map name="headmenu">
					<!--Y.TAKAKUWA Del-S 2015-01-27-->
					<!--<area href="index_en.asp"  target="_top" shape="rect"  coords="0,0,36,13">-->
					<!--Y.TAKAKUWA Del-S 2015-01-27-->
					<area href="userchk.asp?link=touroku/index.html" target="new_window" shape="rect"  coords="42,0,154,13">
					<area href="info.html" shape="rect"  coords="161,0,261,13">
					<area href="JavaScript:openwin()" shape="rect"  coords="267,0,354,13">
					<area href="http://www.hits-h.com/request.asp" shape="rect"  coords="362,0,411,13">
				</map>
			</ul>
		</div>
		
		<div id="left_block">
			<div id="topics">
				<!-- 2014-02-26 Y.TAKAKUWA Upd-S-->
				<!-- <div id="marquee" align="center">      
					<iframe src="denbun.asp" height="20" width="265" scrolling="no" frameborder="0" name="denbun"></iframe>
				</div> -->
				<div id="marquee" align="center" style="position:relative;height:20px;overflow:hidden;">
					<!--[if lt IE 11]>			
					<iframe src="denbun.asp" height="20" width="265" scrolling="no" frameborder="0" name="denbun"></iframe>
					<![endif]-->
					<iframe src="denbun.html" height="20" width="265" scrolling="no" frameborder="0" name="denbun"></iframe>
				</div>
				<!-- 2014-02-26 Y.TAKAKUWA Upd-E-->

				<div id="topic_text" style="width:270px;overflow-x:auto;">
					<table border="0" cellspacing="0" cellpadding="0" width="250">							
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
					</table> 
				</div>
			</div>
	
			<div id="links">
				<ul id="btn">
					<li><a href="http://www.jphkt.co.jp/" target="new_window"><img src="newsite/img/LINK_1.png" alt="�����`�^" /></a></li>
					<li><a href="http://www.sogo-unyu.co.jp/" target="new_window"><img src="newsite/img/LINK_2.png" alt="���݉^�A" /></a></li>
					<li><a href="http://www.nittsu.co.jp/" target="new_window"><img src="newsite/img/LINK_3.png" alt="���{�ʉ^" /></a></li>
					<li><a href="http://www.geneq.co.jp/" target="new_window"><img src="newsite/img/LINK_4.png" alt="�W�F�l�b�N" /></a></li>
					<li><a href="http://www.kamigumi.co.jp" target="new_window"><img src="newsite/img/LINK_5.png" alt="��g" /></a></li>
					<li><a href="http://www.mitsubishi-logistics.co.jp/" target="new_window"><img src="newsite/img/LINK_6.png" alt="�O�H�q��" /></a></li>
					<li><a href="http://www.hakatako-futo.co.jp/" target="new_window"><img src="newsite/img/LINK_7.png" alt="�����`�ӓ�" /></a></li>
					<li><a href="http://port-of-hakata.city.fukuoka.lg.jp/" target="new_window"><img src="newsite/img/LINK_8.png" alt="�����s�`�p��" /></a></li>
					<li><a href="http://www.gct.com.cn/" target="new_window"><img src="newsite/img/LINK_9.png" alt="GUANGZHOU CONTAINER TERMINAL" /></a></li>
					<li><a href="http://www.gnict.com/cn/default.jsp" target="new_window"><img src="newsite/img/LINK_10.png" alt="NANSHA STEVEDORING" /></a></li>
					<li><a href="http://www.sctcn.com/" target="new_window"><img src="newsite/img/LINK_11.png" alt="SHEKOU CONTAINER TERMINALS" /></a></li>
					<li><a href="http://www.cwcct.com//cct/cct_en/publicinf/main/index.aspx" target="new_window"><img src="newsite/img/LINK_12.png" alt="CHIWAN CONTAINER TERMINAL" /></a></li>
					<li><a href="http://www.port.co.th/sitenew/en/" target="new_window"><img src="newsite/img/LINK_15.png" alt="Port Authority of Thailand" /></a></li>
					<li><a href="http://www.tpct.com.tw/" target="new_window"><img src="newsite/img/LINK_14.png" alt="TAIPEI PORT CONTAINER TERMINAL" /></a></li>
					<!-- <li><img src="newsite/img/LINK_btn.png" alt="Blank" /></li>-->
				</ul>
				<div id="linkbanner_left">
					<ul id="bnr">
						<li id="nowfas"><a href="http://www.mlit.go.jp/kowan/nowphas/" target="_blank"><img src="newsite/img/nowfas.jpg" alt="�i�E�t�@�X" /></a></li>
						<li id="mitsui">
								<a href="http://www.mes.co.jp" target="_blank"><img src="newsite/img/mitsui_zosen.jpg" alt="�O�䑢�D" /></a>
								<img src="images/mitsui.gif">
						</li>
					</ul>
				</div>
			</div>
		</div>
   
   		<div id="center_block">
			<p class="title"><img src="images/Explain_title.gif" title="" border="0"></p>
			<p class="description"><img src="images/Explain.gif" title="" border="0"></p>

			<div id="global_menu">
				<ul>
					<li id="gm_exinfo"><a href="userchk.asp?link=expentry.asp" target="_top"><span>�A�o�R���e�i���Ɖ�</span></a></li>
					<li id="gm_entry"><a href="userchk.asp?link=predef/dmi000F.asp" target="_top"><span>���O������</span></a></li>
					<li id="gm_iminfo"><a href="userchk.asp?link=impentry.asp" target="_top"><span>�A���R���e�i���Ɖ�</span></a></li>
					<!-- '2010.5.17 Mod-S MES aoyagi �V���g���@�\�̋x�~���\���̏ꍇ�́ushuttle/shuttle-rest.html�v�������N���� -->
					<!-- <li id="gm_reservation"><a href="userchk.asp?link=Shuttle/SYWB013.asp" target="_top"><span>�V���g���\��</span></a></li> -->
					<li id="gm_reservation"><a href="shuttle/shuttle-rest.html" target="_top"><span>�V���g���\��</span></a></li>
					<!-- '2010.5.17 Mod-E MES aoyagi -->
					<li id="gm_adinfo"><a href="userchk.asp?link=arvdepinfo.asp" target="_top"><span>�����ݏ��Ɖ�</span></a></li>
					<li id="gm_request"><a href="userchk.asp?link=SendStatus/sst000F.asp" target="_top"><span>�A���X�e�[�^�X�z�M�˗�</span></a></li>
					<li id="gm_cyinfo"><a href="userchk.asp?link=terminal.asp" target="_top"><span>CY���G�󋵏Ɖ�</span></a></li>
					<li id="gm_qa"><a href="../qa/index.html" target="_top"><span>�p���`</span></a></li>
				</ul>
			</div>
		</div>
		
		<div id="right_block">
			<script type="text/javascript" language="javascript">
				// flash({ src : 'newsite/swf/top.mp4',
				// 	w   : 259,
				// 	h   : 419 });
			</script>
			<!-- <noscript>
				<video width="259" height="419" autoplay loop id="vFlash">
					<source src="newsite/swf/top.webm" type="video/webm">
				</video>
			
				<object type="application/x-shockwave-flash" data="newsite/swf/top.swf" width="259" height="419"><param name="movie" value="newsite/swf/top.swf" /></object>
			</noscript> -->
			<!-- <img src="newsite/swf/top.gif" /> -->
			<!-- <script type="text/javascript" src="newsite/swf/top.min.js"></script>
			<script type="text/javascript" src="newsite/swf/sprite.js"></script> -->

			<!-- // Edited by AK.DELAROSA 2021/01/11 start -->
			<img src="newsite/scripts/top.gif" id="gif_alternative" width="259" height="419" style="position: absolute; display: none;"/>

			<div id="animation_container" style="width:259px; height:419px; display: none;">
				<canvas id="canvas" width="259" height="419" style="position: absolute; display: block;"></canvas>
				<div id="dom_overlay_container" style="pointer-events:none; overflow:hidden; width:259px; height:419px; position: absolute; left: 0px; top: 0px; display: block;">
				</div>
			</div>

			<!-- // Edited by AK.DELAROSA 2021/01/11 End -->
		</div>
		
		<div id="linkbanner">
			<div id="zentai"> <!--��������-->
				<!-- 2015/07/02 mod by MES �o�i�[�̌����擾�@�\�̒ǉ� -->
				<a href="bannerlog.asp?longid=l101&logno=01&linkname=�o�i�[�P&linkurl=http://www.hakatako-futo.co.jp/" target="_blank"><img src="images/�o�i�[2011.gif" title="" id="gazou1" border=0  width="130" height="60"></a><!--�摜�P�`�R�̐ݒ�-->
				
				<!--2012.4.24 Mod-S MES Aoyagi �o�i�[�X�V �u�L����W���v����u�D�����Ì����v-->
				<a href="bannerlog.asp?longid=l102&logno=01&linkname=�o�i�[�Q&linkurl=http://www.ocean-commerce.co.jp/hakata/" target="_blank" style="color:white"><img src="images/futo-kensaku.gif" title="" id="gazou2"  border=0 width="130" height="60"></a>
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="gazou2" border=0 ></a> -->
				<!-- 2012.4.24 Mod-E MES Aoyagi �o�i�[�X�V -->
				<!-- 2014.5.7 Mod-S MES Aoyagi -->
				<!-- <a href="http://www.unicarriers.co.jp/" target="_blank"><img src="images/banner.gif" title="" id="gazou3"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l103&logno=01&linkname=�o�i�[�R&linkurl=http://www.unicarriers.co.jp/" target="_blank"><img src="images/UNICARRIERS.gif" title="" id="gazou3"  border=0 ></a>
				<!-- 2014.5.7 Mod-S MES Aoyagi -->
				<!-- 2016.7.15 Mod-S MES Aoyagi �o�i�[�X�V �u�L����W���v����u���M�C�^�v�a-->
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="gazou4"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l104&logno=01&linkname=�o�i�[�S&linkurl=http://www.seihou.jp" target="_blank" style="color:white"><img src="images/seihou.jpg" title="" id="gazou4"  border=0 width="130" height="60" alt=""></a> 
				<!--2016.7.15 Mod-E MES Aoyagi �o�i�[�X�V �u�L����W���v����u���M�C�^�v�a-->
				<a href="bannerlog.asp?longid=l105&logno=01&linkname=�o�i�[�T&linkurl=http://www.seiko-denki.co.jp/" target="_blank"><img src="images/SEIKO.JPG" title="" id="gazou5"  border=0 ></a>
				<!--2015.5.15 Mod-S Cosmo Nogami �o�i�[�X�V GENEQ-->		
				<!-- <a><img src="images/Blankbanner3.JPG" title="" id="gazou6" border=0 ></a> -->
				<a href="bannerlog.asp?longid=l106&logno=01&linkname=�o�i�[�U&linkurl=http://www.geneq.co.jp/" target="_blank"><img src="images/GENEQ_cm.gif" title="" id="gazou6"  border=0 ></a>
				<!--2015.5.15 Mod-E Cosmo Nogami �o�i�[�X�V GENEQ-->		
				<!-- 2014.5.7 Mod-S MES Aoyagi -->	
				<!-- <a href="http://www.idex.co.jp" target="_blank"><img src="images/Image00002.jpg" title="" id="gazou7"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l107&logno=01&linkname=�o�i�[�V&linkurl=http://www.idex.co.jp/" target="_blank"><img src="images/Image00002.gif" title="" id="gazou7"  border=0 ></a>
				<!-- 2014.5.7 Mod-E MES Aoyagi -->	
				<!--2016.10.14 Mod-S MES Aoyagi �o�i�[�X�V �uLGX�v����u���؎������v�a-->		
				<!-- <a href="bannerlog.asp?longid=l108&logno=01&linkname=�o�i�[�W&linkurl=http://www.ditp.go.th/japan/download/article/article_20160722170749.pdf" target="_blank"><img src="images/LGX16.jpg" title="" id="gazou8"  border=0 ></a> -->
				<a><img src="images/Blankbanner4.JPG" title="" id="gazou8"  border=0 ></a>
				<!--2016.10.14 Mod-E MES Aoyagi �o�i�[�X�V -->	

		
				<a href="bannerlog.asp?longid=l101&logno=01&linkname=�o�i�[�P&linkurl=http://www.hakatako-futo.co.jp/" target="_blank"><img src="images/�o�i�[2011.gif" title="" id="b_gazou1" border=0  width="130" height="60"></a><!--�摜�P�`�R�̐ݒ�-->
				<!--2012.4.24 Mod-S MES Aoyagi �o�i�[�X�V �u�L����W���v����u�D�����Ì����v-->
				<a href="bannerlog.asp?longid=l102&logno=01&linkname=�o�i�[�Q&linkurl=http://www.ocean-commerce.co.jp/hakata/" target="_blank" style="color:white"><img src="images/futo-kensaku.gif" title="" id="b_gazou2"  border=0 width="130" height="60"></a>
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="b_gazou2" border=0 ></a> -->
				<!-- 2012.4.24 Mod-E MES Aoyagi �o�i�[�X�V -->	
				<!-- 2014.5.7 Mod-S MES Aoyagi -->
				<!-- <a href="http://www.unicarriers.co.jp/" target="_blank"><img src="images/banner.gif" title="" id="b_gazou3"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l103&logno=01&linkname=�o�i�[�R&linkurl=http://www.unicarriers.co.jp/" target="_blank"><img src="images/UNICARRIERS.gif" title="" id="b_gazou3"  border=0 ></a> 
				<!-- 2014.5.7 Mod-E MES Aoyagi -->
				<!--2016.7.15 Mod-S MES Aoyagi �o�i�[�X�V �u�L����W���v����u���M�C�^�v�a-->
				<!-- <a><img src="images/Blankbanner1.JPG" title="" id="b_gazou4"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l104&logno=01&linkname=�o�i�[�S&linkurl=http://www.seihou.jp" target="_blank" style="color:white"><img src="images/seihou.jpg" title="" id="b_gazou4"  border=0 width="130" height="60" alt=""></a> 
				<!-- 2016.7.15 Mod-E MES Aoyagi �o�i�[�X�V �u�L����W���v����u���M�C�^�v�a-->
				<a href="bannerlog.asp?longid=l105&logno=01&linkname=�o�i�[�T&linkurl=http://www.seiko-denki.co.jp/" target="_blank"><img src="images/SEIKO.JPG" title="" id="b_gazou5"  border=0 ></a>
				<!-- 2015.5.15 Mod-S Cosmo Nogami �o�i�[�X�V GENEQ-->		
				<!-- <a><img src="images/Blankbanner3.JPG" title="" id="b_gazou6" border=0 ></a> -->
				<a href="bannerlog.asp?longid=l106&logno=01&linkname=�o�i�[�U&linkurl=http://www.geneq.co.jp/" target="_blank"><img src="images/GENEQ_cm.gif" title="" id="b_gazou6"  border=0 ></a>
				<!-- 2015.5.15 Mod-E Cosmo Nogami �o�i�[�X�V GENEQ-->		
				<!-- 2014.5.7 Mod-S MES Aoyagi -->		
				<!-- <a href="http://www.idex.co.jp" target="_blank"><img src="images/Image00002.jpg" title="" id="b_gazou7"  border=0 ></a> -->
				<a href="bannerlog.asp?longid=l107&logno=01&linkname=�o�i�[�V&linkurl=http://www.idex.co.jp/" target="_blank"><img src="images/Image00002.gif" title="" id="b_gazou7"  border=0 ></a>
				<!-- 2014.5.7 Mod-E MES Aoyagi -->
				<!-- 2016.10.14 Mod-S MES Aoyagi �o�i�[�X�V �uLGX�v����u���؎������v�a-->		
				<!-- <a href="bannerlog.asp?longid=l108&logno=01&linkname=�o�i�[�W&linkurl=http://www.ditp.go.th/japan/download/article/article_20160722170749.pdf" target="_blank"><img src="images/LGX16.jpg" title="" id="b_gazou8"  border=0 ></a> -->
				<a><img src="images/Blankbanner4.JPG" title="" id="b_gazou8"  border=0 ></a> 
				<!-- 2016.10.14 Mod-E MES Aoyagi �o�i�[�X�V -->
			</div><!--�����܂�-->
			<!-- <img src="newsite/img/link_dummy.jpg" width="980" height="59" alt="�����N�o�i�[�̃_�~�[�ł�" />	-->
   		</div>
		   
		<div id="footer">
			<% '2013/09/27 Mw.Tanaka Add-S %>
			<img src="newsite/img/footer.jpg" border="0" usemap="#footerlink">
			<map name="footerlink">
				<area href="http://www.hits-h.com/ija/" target="_top" shape="rect"  coords="95,50,250,70">
				<area href="http://www.hits-h.com/sp/index.aspx" target="_top" shape="rect" coords="330,50,520,70">
				<% '2017/02/21 MES Add-S %>
				<area href="appinfo.asp" target="_blink" shape="rect" coords="330,100,450,115">
				<% '2017/02/21 MES Add-E %>
			</map>
			<% '2013/09/27 Mw.Tanaka Add-E %>
			<div id="get_flash">
				<a href="http://get.adobe.com/jp/flashplayer/" target="_blank"><img src="newsite/img/get_flashplayer.jpg" width="66" height="16" alt="get_flashplayer" /></a>
				<p><img src="images/Get_flash.gif" title=""border=0></p>
			</div>
			<div id="get_adobe_reader">
				<a href="http://get.adobe.com/jp/reader/" target="_blank"><img src="newsite/img/get_adobe_reader.jpg" width="58" height="16" alt="get_adobe_reader" /></a>
				<p><img src="images/Get_reader.gif" title=""border=0 ></p>
			</div>
			<p id="copyright">Copyright(c) 2010 Hakata Port Terminal Co., Ltd. All Rights Reserved.</p>
   		</div>
  	</div>

	
</body>
</html>
