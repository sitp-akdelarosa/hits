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
	rel_path="/" & Replace(Right(param(0),len(param(0))-len(home_path)),"\","/")

	i=0
	'''file_info�z��Ƀt�@�C���̍쐬���Ɩ��O���i�[
	For Each fil In fic
		if(DateDiff("d",fil.DateLastModified,Date)<=CInt(param(1))) then '''�����|�쐬��<=�������
			file_info(i)=Left(fil.DateLastModified,4) & "�N" & Mid(fil.DateLastModified,6,2) & "��" & Mid(fil.DateLastModified,9,2) & "��" & ":" & fil.Name & ":1"
		else		'''�����|�쐬��>�������
			file_info(i)=Left(fil.DateLastModified,4) & "�N" & Mid(fil.DateLastModified,6,2) & "��" & Mid(fil.DateLastModified,9,2) & "��" & ":" & fil.Name & ":0"
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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ja" lang="ja">

 <head>
  <meta http-equiv="Content-Type" content="text/html ; charset=Shift_JIS" />
  <meta http-equiv="Content-Style-Type" content="text/css" />
  <meta http-equiv="Content-Script-Type" content="text/javascript" />
  <title>�����`����IT�V�X�e��</title>
  <meta name="description" content="�����������ĉ�����" />
  <meta name="keywords" content="�L�[���[�h�����ĉ�����" />
  <meta name="author" content="����҂����ĉ�����" />
  <link rel="stylesheet" href="../newsite/css/main.css" type="text/css" />
  <script src="../newsite/scripts/flash.js" type="text/javascript"></script>  
 </head>


 <body>


  <div id="main-container">

   <div id="header">
    <ul id="top_navi">
     <li><a href="http://www.hits-h.com/English/info.html" target="_top">Terms of Service</a></li>
     <li>|</li>     
     <li><a href="index.asp" target="_top">Japanese</a></li>
    </ul>
   </div>
   
   <div id="left_block">
    <div id="topics">
     <div id="marquee" align="center">      
		<iframe src="denbun.asp" height="21" width="265" scrolling="no" frameborder="0" name="denbun"></iframe>
     </div>
     <div id="topic_text">
	 	<Table border="0" cellspacing="0" cellpadding="0" width="250">							
		<% 
			If file_num >0 then 
				For k=0 to file_num-1
					file_data=split(file_info(k),":")
					j=0
					for each fd in file_data
						f(k,j)=fd
						j=j+1
					next
					response.write "<tr>"
					if f(k,2)=1 then
						response.write "<td width='10'>"
						response.write "<img src='../inform/images/new2.gif' border='0'>"
						response.write "</td>"					
					end if
					if f(k,2)<>1 then
						response.write "<td colspan='2' width='100%'>"
					else
						response.write "<td width='100%'>"
					end if
					
					response.write "<p>"
					response.write "<a href=" & rel_path & f(k,1) & " target='_blank'>" & f(k,0) & " " & left(f(k,1),len(f(k,1))-4) & "</a>"
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
       <li><a href="http://www.jphkt.co.jp/" target="new_window"><img src="../newsite/img/LINK_1.png" alt="�����`�^" /></a></li>
       <li><a href="http://www.sogo-unyu.co.jp/" target="new_window"><img src="../newsite/img/LINK_2.png" alt="���݉^�A" /></a></li>
       <li><a href="http://www.nittsu.co.jp/" target="new_window"><img src="../newsite/img/LINK_3.png" alt="���{�ʉ^" /></a></li>
       <li><a href="http://www.geneq.co.jp/" target="new_window"><img src="../newsite/img/LINK_4.png" alt="�W�F�l�b�N" /></a></li>
       <li><img src="../newsite/img/LINK_5.png" alt="��g" /></li>
       <li><img src="../newsite/img/LINK_6.png" alt="�O�H�q��" /></li>
       <li><a href="http://www.hakatako-futo.co.jp/" target="new_window"><img src="../newsite/img/LINK_7.png" alt="�����`�ӓ�" /></a></li>
       <li><a href="http://www.port-of-hakata.or.jp/" target="new_window"><img src="../newsite/img/LINK_8.png" alt="�����s�`�p��" /></a></li>
       <li><img src="../newsite/img/LINK_9.png" alt="GUANGZHOU CONTAINER TERMINAL" /></li>
       <li><img src="../newsite/img/LINK_10.png" alt="NANSHA STEVEDORING" /></li>
       <li><a href="http://www.sctcn.com/english/default.aspx" target="new_window"><img src="../newsite/img/LINK_11.png" alt="SHEKOU CONTAINER TERMINALS" /></a></li>
       <li><a href="http://www.cwcct.com//cct/cct_en/publicinf/main/index.aspx" target="new_window"><img src="../newsite/img/LINK_12.png" alt="CHIWAN CONTAINER TERMINAL" /></a></li>
       <li><img src="../newsite/img/LINK_btn.png" alt="Blank" /></li>
       <li><img src="../newsite/img/LINK_btn.png" alt="Blank" /></li>
     </ul>
    <div id="linkbanner_left">
     <ul id="bnr">
       <li id="nowfas"><a href="http://www.mlit.go.jp/kowan/nowphas/" target="_blank"><img src="../newsite/img/nowfas.jpg" alt="�i�E�t�@�X" /></a></li>
       <li id="mitsui"><a href="http://www.mes.co.jp" target="_blank"><img src="../newsite/img/mitsui_zosen.jpg" alt="�O�䑢�D" /></a><br /><p>MITSUI ENGINEERING &amp;  SHIPBUILDING CO.,LTD.</p></li>
     </ul>
    </div>
    </div>
    
   </div>
   
   <div id="center_block">
    <p class="title">�����I�����̎������T�|�[�g����IT�V�X�e��</p>
    <p class="description">HiTS�́A�C���^�[�l�b�g�𗘗p���āu�׎�v�u�C�݁v�u���^�v<br />�u�`�^���Ǝҁv�Ȃǂ̊Ԃł���肷��A�o���Ɩ��̎w���`�B���A<br />PC��g�ѓd�b�ɂ���čs���閳���V�X�e���ł��B</p>
    <div id="global_menu">
     <ul>
     <li id="gm_exinfo_en"><a href="http://www.hits-h.com/English/expentry.asp" target="_top"><span>EXPORT CONTAINER INFORMATION</span></a></li>
     <li><img src="../newsite/img/global_menu_off.png" alt="Blank" /></li>
     <li id="gm_iminfo_en"><a href="http://www.hits-h.com/English/impentry.asp" target="_top"><span>IMPORT CONTAINER INFORMATION</span></a></li>
     <li><img src="../newsite/img/global_menu_off.png" alt="Blank" /></li>
     <li><img src="../newsite/img/global_menu_off.png" alt="Blank" /></li>
     <li><img src="../newsite/img/global_menu_off.png" alt="Blank" /></li>
     <li><img src="../newsite/img/global_menu_off.png" alt="Blank" /></li>
     <li><img src="../newsite/img/global_menu_off.png" alt="Blank" /></li>
     </ul>
    </div>
   </div>

   <div id="right_block">
    <script type="text/javascript" language="javascript"><!--

flash({ src : 'swf/top.swf',
        w   : 259,
	h   : 419 });
//-->
</script>
<noscript><object type="application/x-shockwave-flash" data="../newsite/swf/top.swf" width="259" height="419"><param name="movie" value="../newsite/swf/top.swf" /></object></noscript>
   </div>
   
   <div id="linkbanner">
    <img src="../newsite/img/link_dummy.jpg" width="980" height="59" alt="�����N�o�i�[�̃_�~�[�ł�" />
   </div>
    
   <div id="footer">
    <div id="get_flash">
     <a href="http://get.adobe.com/jp/flashplayer/" target="_blank"><img src="../newsite/img/get_flashplayer.jpg" width="66" height="16" alt="get_flashplayer" /></a>
      <p>���悪�\������Ȃ��ꍇ�́A<br />���̃A�C�R���̃����N�悩��AdobeFlashPlayer���_�E�����[�h���Ă��������B</p>
    </div>
    <div id="get_adobe_reader">
     <a href="http://get.adobe.com/jp/reader/" target="_blank"><img src="../newsite/img/get_adobe_reader.jpg" width="58" height="16" alt="get_adobe_reader" /></a>
      <p>PDF���\������Ȃ��ꍇ�́A<br />���̃A�C�R���̃����N�悩��AdobeReader���_�E�����[�h���Ă��������B</p>
    </div>
    <p id="copyright">Copyright(c) 2010 Hakata Port Terminal Co., Ltd. All Rights Reserved.</p>
   </div>
   
  </div>

 </body>


</html>
