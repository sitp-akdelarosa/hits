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
	rel_path="/" & Replace(Right(param(0),len(param(0))-len(home_path)),"\","/")

	i=0
	'''file_info配列にファイルの作成日と名前を格納
	For Each fil In fic
		if(DateDiff("d",fil.DateLastModified,Date)<=CInt(param(1))) then '''今日－作成日<=特定期間
			file_info(i)=Left(fil.DateLastModified,4) & "年" & Mid(fil.DateLastModified,6,2) & "月" & Mid(fil.DateLastModified,9,2) & "日" & "|" & Mid(fil.DateLastModified,12,8) & "|" & fil.Name & "|1"
		else		'''今日－作成日>特定期間
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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ja" lang="ja">

 <head>
  <meta http-equiv="Content-Type" content="text/html ; charset=Shift_JIS" />
  <meta http-equiv="Content-Style-Type" content="text/css" />
  <meta http-equiv="Content-Script-Type" content="text/javascript" />
  <title>博多港物流ITシステム</title>
  <meta name="description" content="説明文を入れて下さい" />
  <meta name="keywords" content="キーワードを入れて下さい" />
  <meta name="author" content="著作者を入れて下さい" />
  <link rel="stylesheet" href="newsite/css/main.css" type="text/css" />
  <script src="newsite/scripts/sc_1.js" type="text/javascript"></script>
  <script src="newsite/scripts/flash.js" type="text/javascript"></script>
  <SCRIPT language=JavaScript>
	function openwin(){
	
		var w=900;
		var h=550;
		var l=0;
		var t=0;
		if(screen.width){
			l=(screen.width-w)/2;
		}
		if(screen.availWidth){
			l=(screen.availWidth-w)/2;
		}
		if(screen.height){
			t=(screen.height-h)/2;
		}
		if(screen.availHeight){
			t=(screen.availHeight-h)/2;
		}
		
		var win=window.open("../download/download_list.asp","","status=no,width="+w+",height="+h+",top="+t+",left="+l);
	}
  </SCRIPT>
 </head>


 <body>


  <div id="main-container">

   <div id="header">
    <ul id="login">
    <a href="userchk.asp?link=index.asp" target="_top"><img src="newsite/img/login_btn.gif" border="0"></a>
    </ul>
    <ul id="top_navi">
    <img src="images/headmenu.gif" border="0" usemap="#headmenu">
    <map name="headmenu">
        <area href="index_en.asp"  target="_top" shape="rect"  coords="0,0,36,13">
        <area href="userchk.asp?link=touroku/index.html" target="new_window" shape="rect"  coords="42,0,154,13">
        <area href="info.html" shape="rect"  coords="161,0,261,13">
        <area href="JavaScript:openwin()" shape="rect"  coords="267,0,354,13">
        <area href="http://www.hits-h.com/request.asp" shape="rect"  coords="362,0,411,13">
    </map>
    </ul>
   </div>
   
   <div id="left_block">
    <div id="topics">
     <div id="marquee" align="center">      
		<iframe src="denbun.asp" height="20" width="265" scrolling="no" frameborder="0" name="denbun"></iframe>
     </div>
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
    </div>
    
    <div id="links">
     <ul id="btn">
       <li><a href="http://www.jphkt.co.jp/" target="new_window"><img src="newsite/img/LINK_1.png" alt="博多港運" /></a></li>
       <li><a href="http://www.sogo-unyu.co.jp/" target="new_window"><img src="newsite/img/LINK_2.png" alt="相互運輸" /></a></li>
       <li><a href="http://www.nittsu.co.jp/" target="new_window"><img src="newsite/img/LINK_3.png" alt="日本通運" /></a></li>
       <li><a href="http://www.geneq.co.jp/" target="new_window"><img src="newsite/img/LINK_4.png" alt="ジェネック" /></a></li>
       <li><img src="newsite/img/LINK_5.png" alt="上組" /></li>
       <li><img src="newsite/img/LINK_6.png" alt="三菱倉庫" /></li>
       <li><a href="http://www.hakatako-futo.co.jp/" target="new_window"><img src="newsite/img/LINK_7.png" alt="博多港ふ頭" /></a></li>
       <li><a href="http://www.port-of-hakata.or.jp/" target="new_window"><img src="newsite/img/LINK_8.png" alt="福岡市港湾局" /></a></li>
       <li><a href="http://www.gct.com.cn/" target="new_window"><img src="newsite/img/LINK_9.png" alt="GUANGZHOU CONTAINER TERMINAL" /></a></li>
       <li><a href="http://www.gnict.com/cn/default.jsp" target="new_window"><img src="newsite/img/LINK_10.png" alt="NANSHA STEVEDORING" /></a></li>
       <li><a href="http://www.sctcn.com/" target="new_window"><img src="newsite/img/LINK_11.png" alt="SHEKOU CONTAINER TERMINALS" /></a></li>
       <li><a href="http://www.cwcct.com//cct/cct_en/publicinf/main/index.aspx" target="new_window"><img src="newsite/img/LINK_12.png" alt="CHIWAN CONTAINER TERMINAL" /></a></li>
       <li><img src="newsite/img/LINK_btn.png" alt="Blank" /></li>
       <li><img src="newsite/img/LINK_btn.png" alt="Blank" /></li>
     </ul>
    <div id="linkbanner_left">
     <ul id="bnr">
       <li id="nowfas"><a href="http://www.mlit.go.jp/kowan/nowphas/" target="_blank"><img src="newsite/img/nowfas.jpg" alt="ナウファス" /></a></li>
       <li id="mitsui">
       		<a href="http://www.mes.co.jp" target="_blank"><img src="newsite/img/mitsui_zosen.jpg" alt="三井造船" /></a>
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
     <li id="gm_exinfo"><a href="userchk.asp?link=expentry.asp" target="_top"><span>輸出コンテナ情報照会</span></a></li>
     <li id="gm_entry"><a href="userchk.asp?link=predef/dmi000F.asp" target="_top"><span>事前情報入力</span></a></li>
     <li id="gm_iminfo"><a href="userchk.asp?link=impentry.asp" target="_top"><span>輸入コンテナ情報照会</span></a></li>
<!-- '2010.5.17 Mod-S MES aoyagi シャトル機能の休止中表示の場合は「shuttle/shuttle-rest.html」をリンクする -->
<!--     <li id="gm_reservation"><a href="userchk.asp?link=Shuttle/SYWB013.asp" target="_top"><span>シャトル予約</span></a></li> -->
     <li id="gm_reservation"><a href="shuttle/shuttle-rest.html" target="_top"><span>シャトル予約</span></a></li>
<!-- '2010.5.17 Mod-E MES aoyagi -->
     <li id="gm_adinfo"><a href="userchk.asp?link=arvdepinfo.asp" target="_top"><span>着離岸情報照会</span></a></li>
     <li id="gm_request"><a href="userchk.asp?link=SendStatus/sst000F.asp" target="_top"><span>輸入ステータス配信依頼</span></a></li>
     <li id="gm_cyinfo"><a href="userchk.asp?link=terminal.asp" target="_top"><span>CY混雑状況照会</span></a></li>
     <li id="gm_qa"><a href="../qa/index.html" target="_top"><span>Ｑ＆Ａ</span></a></li>
     </ul>
    </div>
   </div>

   <div id="right_block">
    <script type="text/javascript" language="javascript"><!--

flash({ src : 'newsite/swf/top.swf',
        w   : 259,
	h   : 419 });
//-->
</script>
<noscript><object type="application/x-shockwave-flash" data="newsite/swf/top.swf" width="259" height="419"><param name="movie" value="newsite/swf/top.swf" /></object></noscript>
   </div>
   
   <div id="linkbanner">
		<div id="zentai"><!--ここから-->

		<a href="http://www.hakatako-futo.co.jp/index.php"><img src="images/バナー2011.gif" title="" id="gazou1" border=0  width="130" height="60"></a><!--画像１～３の設定-->
<!--2010.6.15 Mod-S MES Aoyagi バナー更新 「広告募集中」から「日鐵」殿-->
<!--		<a><img src="images/Blankbanner1.JPG" title="" id="gazou2" border=0 ></a> -->
		<a href="http://ntc.ntsysco.co.jp/"><img src="images/nittetsu.gif" title="" id="gazou2"  border=0 ></a>
<!--2010.6.15 Mod-E MES Aoyagi バナー更新 -->	
		<a href="http://www.tcm.co.jp"><img src="images/banner.gif" title="" id="gazou3"  border=0 ></a>
		<a><img src="images/Blankbanner1.JPG" title="" id="gazou4"  border=0 ></a>
		<a href="http://www.seiko-denki.co.jp"><img src="images/SEIKO.JPG" title="" id="gazou5"  border=0 ></a>
<!--2010.5.17 Mod-S MES Aoyagi バナー更新 「広告募集中」から「鶴丸海運」殿-->		
<!--2011/04/01 Upd-S Fujiyama バナー更新 「広告募集中」から「鶴丸海運」殿 から「広告募集中」-->
		<a><img src="images/Blankbanner3.JPG" title="" id="gazou6" border=0 ></a>
<!--		<a href="http://www.tsurumaru.co.jp"><img src="images/tsurumaru.gif" title="" id="gazou6"  border=0 ></a> -->
<!--2011/04/01 Upd-S Fujiyama-->
<!--2010.5.17 Mod-E MES Aoyagi バナー更新 -->		
		<a href="http://www.idex.co.jp"><img src="images/Image00002.jpg" title="" id="gazou7"  border=0 ></a>
		<a><img src="images/Blankbanner4.JPG" title="" id="gazou8"  border=0 ></a>
		
		
		<a href="http://www.hakatako-futo.co.jp/index.php"><img src="images/バナー2011.gif" title="" id="b_gazou1" border=0  width="130" height="60"></a><!--画像１～３の設定-->
<!--2010.6.15 Mod-S MES Aoyagi バナー更新 「広告募集中」から「日鐵」殿-->
<!--		<a><img src="images/Blankbanner1.JPG" title="" id="b_gazou2" border=0 ></a> -->
		<a href="http://ntc.ntsysco.co.jp/"><img src="images/nittetsu.gif" title="" id="b_gazou2"  border=0 ></a>
<!--2010.6.15 Mod-E MES Aoyagi バナー更新 -->	
		<a href="http://www.tcm.co.jp"><img src="images/banner.gif" title="" id="b_gazou3"  border=0 ></a>
		<a><img src="images/Blankbanner1.JPG" title="" id="b_gazou4"  border=0 ></a>
		<a href="http://www.seiko-denki.co.jp"><img src="images/SEIKO.JPG" title="" id="b_gazou5"  border=0 ></a>
<!--2010.5.17 Mod-S MES Aoyagi バナー更新 「広告募集中」から「鶴丸海運」殿-->		
<!--2011/03/31 Upd-S Fujiyama バナー更新 「広告募集中」から「鶴丸海運」殿 から「広告募集中」-->
		<a><img src="images/Blankbanner1.JPG" title="" id="b_gazou6" border=0 ></a>
<!--		<a href="http://www.tsurumaru.co.jp"><img src="images/tsurumaru.gif" title="" id="b_gazou6"  border=0 ></a> -->
<!--2011/03/31 Upd-S Fujiyama-->
<!--2010.5.17 Mod-E MES Aoyagi バナー更新 -->		
		<a href="http://www.idex.co.jp"><img src="images/Image00002.jpg" title="" id="b_gazou7"  border=0 ></a>
		<a><img src="images/Blankbanner4.JPG" title="" id="b_gazou8"  border=0 ></a>
		
		</div><!--ここまで-->
<!--    <img src="newsite/img/link_dummy.jpg" width="980" height="59" alt="リンクバナーのダミーです" />	-->
   </div>
    
   <div id="footer">
   
   
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
