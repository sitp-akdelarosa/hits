/********************************************************************
 * �J�����_�[�ɂ����t���̓X�N���v�g 20021115
 *
 * ( ���L�X�N���v�g�͉������\�ł����܂����������炸�ɂ��̂܂܃y�[�X
 *   �g���邾���ł������p����������悤�ɏ����Ă���܂� )
 *
 Syntax : 

    wrtCalendarLayLay( formElementObject , event , ���t�^�C�v ) 


  �g������INPUT���̓^�O��onFocus="wrtCalendarLay(this,event)"���y�[
  �X�g���܂��BFORM��INPUT�^�O�ɈႤ���O(NAME����)��Y�ꂸ�ɕt���Ă��� 
  �Ă��������B

  ���� : 

     formElementObject  ���͂������t�H�[���G�������g 

     event  �C�x���g( event | null ) 
               event�ŃJ�[�\���̂��΂Ɍ����Anull�ŌŒ� 

     ���t�^�C�v   
     
               'yyyy' �� 2005 
               'yyyy/mm' �� 2005/3 
               'yyyy/mm/dd' �� 2002/2/19 
               'mm/dd' �� 12/24 
               'mm' �� 3 
               'dd' �� 31 
               'yyyy/mm/dd[�j]' �� 2002/6/4 [��] 
               'yyyy/mm/dd(�j)' �� 2002/6/4 (��) 
               'yyyy�Nmm��dd��(�j)'�� 2002�N2��19��(��) 
               'mm��dd��' �� 1��1�� 
               'mm��dd��(�j)' �� 1��1��(��) 

               *�@�f�t�H���g�́A'yyyy/mm/dd' 
               *  Mac��IE�ł͊����̍�����^�C�v�͎g���܂���B

  ��1 : wrtCalendarLay( this , event ) 
  ��2 : wrtCalendarLay( this , event , 'mm/dd' ) 
  ��3 : wrtCalendarLay( document.form1.element3 , event , 'mm��dd��(�j)' ) 
  ��4 : wrtCalendarLay( this.form.element['e0'] , event , 'yyyy�Nmm��dd��(�j)' ) 
  
  Example : 
  
      ��t��1:<input name="e1" type="text" 
                   onFocus="wrtCalendarLay(this,event)"> 
  
      ��t��2:<input name="e2" type="text" 
                   onFocus="wrtCalendarLay(this,event,'yyyy�Nmm��dd��(�j)')"> 
 */

  var now    = new Date()
  var absnow = now
  var Win    = navigator.userAgent.indexOf('Win')!=-1
  var Mac    = navigator.userAgent.indexOf('Mac')!=-1
  var X11    = navigator.userAgent.indexOf('X11')!=-1
  var Moz    = navigator.userAgent.indexOf('Gecko')!=-1
  var msie   = navigator.userAgent.indexOf('MSIE')!=-1
  var bwlang = getBrowserLANG()
  var _utf   = "��".length > 1 
  var nonja  = ( _utf || bwlang == 'en')
  if( nonja )
    var week   = new Array('sun','mon','tue','wed','thu','fri','sat');
  else 
    var week   = new Array('��','��','��','��','��','��','�y');

  //���͌�ޔ��ʒu
  if( Mac && msie ){ var gox=2000 ; var goy=2000 }
  else             { var gox=-300 ; var goy=-300 }
  //n4�p���C���[�o�͈ʒu
  if(document.layers){var n4_left=100 ; var n4_top= 100 }

  calendarLay['calendar']=new calendarLay('calendar',-100,-100,'')

  function wrtCalendarLay(oj,e,dateType,arg1){
  
    set_event__wrtCalendarLay() //�C�x���g�L���v�`���[�X�^�[�g

    // ���t�^�C�v�f�t�H���g�l�ݒ�Ƌ󔒕����񏜋�
    if(!arguments[2])dateType='yyyy/mm/dd';
    else arguments[2].split(' ').join('').split('�@').join('')

    // ���ړ��t���O�f�t�H���g�ݒ�
    if(!arguments[3])arg1=0

    wrtCalendarLay.arg1=arg1
    wrtCalendarLay.oj=oj
    wrtCalendarLay.dateType=dateType
  
    // ���ݏ�����
    if(arg1==0)now = new Date()
  
    // �N�����擾
    nowdate  = now.getDate()
    nowmonth = now.getMonth()
    nowyear  = now.getYear()
  
    // ���ړ�����
    if(nowmonth==11 && arg1 > 0){        //12����arg1��+�Ȃ�
      nowmonth = -1 + arg1 ; nowyear++   //����arg1-1;1�N���Z
    } else if(nowmonth==0 && arg1 < 0){  //1����arg1��-�Ȃ�
      nowmonth = 12 + arg1 ; nowyear--   //����arg1+12;1�N���Z
    } else {
      nowmonth +=  arg1                  //2-11���Ȃ猎��+arg1
    }
  
    // 2000�N���Ή�
    if(nowyear<1900)nowyear=1900+nowyear
  
    // ���݌����m��
    now   = new Date(nowyear,nowmonth,1)
  
    // YYYYMM�쐬
    nowyyyymm=nowyear*100+nowmonth
  
    // YYYY/MM�쐬
    nowtitleyyyymm=nowyear+'/'+(nowmonth + 1)
  
    // �J�����_�[�\�z�p����̎擾
    fstday   = now                                           //������1��
    startday = fstday - ( fstday.getDay() * 1000*60*60*24 )  //�ŏ��̓��j��
    startday = new Date(startday)
  
    // �J�����_�[�\�z�pHTML
    ddata = ''
    ddata += '<FORM>\n'
    ddata += '<TABLE BORDER=0 BGCOLOR="#dddddd"  BORDERCOLOR="#dddddd" WIDTH=140 HEIGHT=140\n'
    ddata += 'STYLE="\n'
    ddata += 'font-family      : Arial;\n'
    ddata += 'font-size        : 14px;\n'
    ddata += 'border-top       : 1px outset #ffffff;\n'
    ddata += 'border-right     : 1px outset #888888;\n'
    ddata += 'border-bottom    : 1px outset #555555;\n'
    ddata += 'border-left      : 1px outset #ffffff;"\n'
    ddata += '>\n'

    // Month
    ddata += '   <TR id="trmonth" BGCOLOR=#6699ff BORDERCOLOR=#6699ff WIDTH=140 HEIGHT=14>\n'
      ddata += '   <TH COLSPAN=7 WIDTH=140 HEIGHT=14 ALIGN="right"><NOBR>\n'
      ddata += '   <FONT SIZE="4" FACE="Arial">\n'
      ddata +=       nowtitleyyyymm
      ddata += '   </FONT>\n'
      ddata += '<INPUT TYPE=button VALUE="<<" \n'
      ddata += 'onClick="wrtCalendarLay(window.document.'+oj.form.name+'.'+oj.name+',null,\''+dateType+'\',-1)"\n'
      ddata += '><INPUT TYPE=button VALUE="o" \n'
      ddata += 'onClick="wrtCalendarLay(window.document.'+oj.form.name+'.'+oj.name+',null,\''+dateType+'\',0)"\n'
      ddata += '><INPUT TYPE=button VALUE=">>" \n'
      ddata += 'onClick="wrtCalendarLay(window.document.'+oj.form.name+'.'+oj.name+',null,\''+dateType+'\',1)">\n'
      ddata += '</NOBR></TH>\n'
    ddata += '   </TR>\n'
  
    // Week
    ddata += '   <TR BGCOLOR=#00cccc WIDTH=140 HEIGHT=14>\n'
  
    for (i=0;i<7;i++){
      ddata += '   <TH WIDTH=14 HEIGHT=14>\n'
      ddata += '   <FONT SIZE="2">\n'
      ddata +=       week[i]
      ddata += '   </FONT>\n'
      ddata += '   </TH>\n'
    }
    ddata += '   </TR>\n'
  
    // Date
    for(j=0;j<6;j++){
      ddata += '   <TR BGCOLOR=#eeeeee>\n'
      for(i=0;i<7;i++){
        nextday     = startday.getTime() + (i * 1000*60*60*24)
        wrtday      = new Date(nextday)
        wrtdate     = wrtday.getDate()
        wrtmonth    = wrtday.getMonth()
        wrtyear     = wrtday.getYear()
        if(wrtyear < 1900) wrtyear = 1900 + wrtyear
        wrtyyyymm   = wrtyear * 100 + wrtmonth
        wrtyyyymmdd = ''+wrtyear +'/'+ (wrtmonth+1) +'/'+wrtdate
        getday      = getWeek(wrtyyyymmdd)
        var outputdate=eval( getDateType(dateType))
        wrtdateA  = '<A HREF="javascript:function v(){'
        wrtdateA += 'document.'+oj.form.name+'.'+oj.name+'.value=(\''+outputdate
        wrtdateA += '\');if(!(Mac&&document.layers))calendarLay[\'calendar\'].moveLAYOJ(getStyleOj(\'calendar\'),'
        wrtdateA += gox+','+goy+');stop_event__wrtCalendarLay()};v()"   >\n'
        wrtdateA += '<FONT COLOR=#000000>\n'
        wrtdateA += wrtdate
        wrtdateA += '</FONT>\n'
        wrtdateA += '</A>\n'
  
        if(wrtyyyymm != nowyyyymm){ 
          ddata += ' <TD BGCOLOR=#cccccc WIDTH=14 HEIGHT=14>\n'
          ddata += wrtdateA
  
        } else if(   wrtdate == absnow.getDate()   
                  && wrtmonth == absnow.getMonth() 
                  && wrtday.getYear() == absnow.getYear()){
          ddata += ' <TD BGCOLOR=#ff99ff WIDTH=14 HEIGHT=14>\n'
          ddata += '<FONT COLOR="#ffffff">'+wrtdateA+'</FONT>\n'
  
        } else {
          ddata += ' <TD WIDTH=14 HEIGHT=14>\n'
          ddata += wrtdateA
        }
        ddata += '   </TD>\n'
      }
      ddata += '   </TR>\n'
  
      startday = new Date(nextday)
      startday = startday.getTime() + (1000*60*60*24)
      startday = new Date(startday)
    }
    // �X�e�[�^�X�s ���t�^�C�v
    ddata += '   <TR>\n'
      ddata += '   <TD COLSPAN=7 ALIGN=center STYLE="font-size:11px">\n'
//       ddata += wrtCalendarLay.dateType
       ddata += ' <INPUT TYPE=button VALUE="close" \n'
       ddata += 'onClick="moveLAYOJ(getStyleOj(\'calendar\'),'+gox+','+goy+')">\n'
      ddata += '   </TD>\n'
    ddata += '   </TR>\n'
  
    ddata += '</TABLE>\n'
    ddata += '</FORM>\n'
    ddata += '</BODY>\n'
    ddata += '</HTML>\n'
  
    calendarLay['calendar'].outputLAYOJ(getLayOj('calendar'),'')//�ꎞ�N���A
    calendarLay['calendar'].outputLAYOJ(getLayOj('calendar'),ddata)

    if(e!=null){
      if(navigator.userAgent.indexOf('Gecko')!=-1){   //n6,m1�p
        var left = e.currentTarget.offsetLeft +510
        var top  = e.currentTarget.offsetTop + 20
      } else {
        var left = getMouseX(e) - 90
        var top  = getMouseY(e) - 110
      }
      if(document.layers){ var left = n4_left ; var top  = n4_top }//n4�C��
      calendarLay['calendar'].moveLAYOJ(getStyleOj('calendar'),left,top)
  
    }
  
  }

  // �j���擾
  function getWeek(date){
    if(arguments.length>0)date=date
    else date=null
    if(  Mac && msie )//MacIE5�p
      week   = new Array('sun','mon','tue','wed','thu','fri','sat');
    var now  = new Date(date) ;
    return week[now.getDay()] ;
  }
  // �o�͓��t�̃f�[�^�^�C�v
  function getDateType(dateType){
      if(nonja || ( Mac && msie )){ //�������\�L�̉��
        if ( dateType == 'yyyy�Nmm��dd��(�j)')  dateType = 'yyyy/mm/dd(�j)'
        else if( dateType == 'mm��dd��')        dateType = 'mm/dd'
        else if( dateType == 'mm��dd��(�j)')    dateType = 'mm/dd(�j)' 
      }
      switch(dateType){
        case 'yyyy'              
: dtate= "''+wrtyear                                                    " ; break ;
        case 'yyyy/mm'           
: dtate= "''+wrtyear +'/'+ (wrtmonth+1)                                 " ; break ;
        case 'yyyy/mm/dd'        
: dtate= "''+wrtyear +'/'+ (wrtmonth+1) +'/'+wrtdate                    " ; break ;
        case 'mm/dd'             
: dtate= "''+              (wrtmonth+1) +'/'+wrtdate                    " ; break ;
        case 'mm'                
: dtate= "''+              (wrtmonth+1)                                 " ; break ;
        case 'dd'                
: dtate= "''+                                wrtdate                    " ; break ;
        case 'yyyy/mm/dd[�j]'    
: dtate= "''+wrtyear +'/'+ (wrtmonth+1) +'/'+wrtdate +' ['+getday +']'  " ; break ;
        case 'yyyy/mm/dd(�j)'    
: dtate= "''+wrtyear +'/'+ (wrtmonth+1) +'/'+wrtdate +' ('+getday +')'  " ; break ;
        case 'mm/dd(�j)'    
: dtate= "''+              (wrtmonth+1) +'/'+wrtdate +' ('+getday +')'  " ; break ;
        case 'yyyy�Nmm��dd��(�j)'
: dtate= "''+wrtyear +'�N'+ (wrtmonth+1)+'��'+wrtdate +'��('+getday +')'" ; break ;
        case 'mm��dd��'          
: dtate= "''+              (wrtmonth+1) +'��'+wrtdate +'��'             " ; break ;
        case 'mm��dd��(�j)'      
: dtate= "''+              (wrtmonth+1) +'��'+wrtdate +'��('+getday +')'" ; break ;
        default                  
: dtate= "''+wrtyear +'/'+ (wrtmonth+1) +'/'+wrtdate                    " ;
      }
    return dtate
  }

  //--���C���[����
  function calendarLay(layName,x,y,dateType){
    this.id      = layName   // �h���b�O�ł���悤�ɂ��郌�C���[��
    this.x       = x         // ����left�ʒu
    this.y       = y         // ����top�ʒu
    this.dateType = dateType // YYYY/MM/DD
    this.day     = new Array()
    if(document.layers)      //n4�p
      this.div='<layer name="'+layName+'" left="'+x+'" top="'+y+'"\n'
              +'       onfocus="clickElement=\''+layName
                                    +'\';mdown_wrtCalendarLay(event);return false">\n'
              +'<a     href="javascript:void(0)"\n'
              +'       onmousedown="clickElement=\''+layName
                                    +'\';mdown_wrtCalendarLay(event);return false">\n'
              + '</a></layer>\n'
    else                     //n4�ȊO�p
      this.div='<div  id="'+layName+'" class="dragLays"\n'
              +'      onmousedown="clickElement=\''+layName
                                    +'\';mdown_wrtCalendarLay(event);return false"\n'
              +'      style="position:absolute;left:'+x+'px;top:'+y+'px">\n'
              + '</div>\n'
    document.write(this.div)
    return 
  }
  calendarLay.prototype.moveLAYOJ   = moveLAYOJ   //���\�b�h��ǉ�����
  calendarLay.prototype.outputLAYOJ = outputLAYOJ //���\�b�h��ǉ�����
  calendarLay.prototype.zindexLAYOJ = zindexLAYOJ //���\�b�h��ǉ�����

  //--���C���[�ړ�
  function moveLAYOJ(oj,x,y){
    if(document.getElementById){  //e5,e6,n6,m1,o6�p
      oj.left = x
      oj.top  = y
    } else if(document.all){      //e4�p
      oj.pixelLeft = x
      oj.pixelTop  = y
    } else if(document.layers)    //n4�p
      oj.moveTo(x,y)
  }
  //--HTML�o��
  function outputLAYOJ(oj,html){
    if(document.getElementById) oj.innerHTML=html  //n6,m1,e5,e6�p
    else if(document.all) oj.innerHTML=html //e4�p
    else if(document.layers)                       //n4�p
       with(oj.document){
         open()
         write(html)
         close()
      }
  }
  //--���s��Z���Wset 
  function zindexLAYOJ(oj,zindex){
    if(document.getElementById) oj.zIndex=zindex  //n6,m1,e5,e6,o6�p
    else if(document.all)       oj.zIndex=zindex  //e4�p
    else if(document.layers)    oj.zIndex=zindex  //n4�p
  }

  //--layName�Ŏw�肵���I�u�W�F�N�g��Ԃ�(�K��onload��Ɏ��s���邱��)
  function getLayOj(layName){  
    if(document.getElementById) 
      return document.getElementById(layName)           //e5,e6,n6,m1,o6�p
    else if(document.all)   return document.all(layName)    //e4�p
    else if(document.layers)return document.layers[layName] //n4�p
  }
  function getStyleOj(clickElement){  
       return (!!document.layers)?getLayOj(clickElement)
                                 :getLayOj(clickElement).style
  }

  //--�}�E�XX���Wget 
  function getMouseX(e){
    if(window.opera)                            //o6�p
        return e.clientX
    else if(document.all)                       //e4,e5,e6�p
        return document.body.scrollLeft+event.clientX
    else if(document.layers||document.getElementById)
        return e.pageX                          //n4,n6,m1�p
  }

  //--�}�E�XY���Wget 
  function getMouseY(e){
    if(window.opera)                            //o6�p
        return e.clientY
    else if(document.all)                       //e4,e5,e6�p
        return document.body.scrollTop+event.clientY
    else if(document.layers||document.getElementById)
        return e.pageY                          //n4,n6,m1�p
  }

  //--���C���|����X���Wget 
  function getLEFT(layName){
    if(document.all)                            //e4,e5,e6,o6�p
      return document.all(layName).style.pixelLeft
    else if(document.getElementById)            //n6,m1�p
      return (document.getElementById(layName).style.left!="")
              ?parseInt(document.getElementById(layName).style.left):""
    else if(document.layers)                    //n4�p
      return document.layers[layName].left 
  }

  //--���C���|���Y���Wget 
  function getTOP(layName){
    if(document.all)                          //e4,e5,e6,o6�p
      return document.all(layName).style.pixelTop
    else if(document.getElementById)          //n6,m1�p
      return (document.getElementById(layName).style.top!="")
              ?parseInt(document.getElementById(layName).style.top):""
    else if(document.layers)                  //n4�p
      return document.layers[layName].top 
  }

  //--�}�E�X�J�[�\���𓮂����������C���[��moveLAYOJ�œ�����
  function mmove_wrtCalendarLay(e) {
    if(!window.clickElement) return
    if (getLayOj(clickElement)) {
       movetoX = getMouseX(e) - offsetX
       movetoY = getMouseY(e) - offsetY
       var oj=getStyleOj(clickElement)
      calendarLay[clickElement].moveLAYOJ(oj,movetoX,movetoY)
      return false
    }
  }

  //--�}�E�X�{�^����������������
  //  ���C���[���̃J�[�\��offset�ʒu�擾
  function mdown_wrtCalendarLay(e) {
    if(navigator.userAgent.indexOf('Gecko')!=-1)   //n6,m1�p
      if(e.currentTarget.className != 'dragLays') return
      else clickElement = e.currentTarget.id
    var selLay = getLayOj(clickElement)
    if (selLay){
        offsetX = getMouseX(e) - getLEFT(selLay.id)
        offsetY = getMouseY(e) - getTOP(selLay.id)
       if(document.layers){
        offsetX = getMouseX(e)+10 ; offsetY = getMouseY(e)+10
       }
    }
    return false
  }

  //--�}�E�X�{�^�����グ�����h���b�O����
  var zcount = 0
  function mup_wrtCalendarLay(e) {
    if(!window.clickElement) return
    if (getLayOj(clickElement)) {
      calendarLay[clickElement].zindexLAYOJ(
        getStyleOj(clickElement),zcount++)
      clickElement=null
    }
  }

  //--�C�x���g�L���v�`���[�J�n
  function set_event__wrtCalendarLay(){
    document.onmousemove = mmove_wrtCalendarLay   //n4,m1,n6,e4,e5,e6,o6�p
    document.onmouseup   = mup_wrtCalendarLay     //n4,m1,n6,e4,e5,e6,o6�p
    if(navigator.userAgent.indexOf('Gecko')!=-1)  //m1,n6�p
      document.onmousedown = mdown_wrtCalendarLay
    if(document.layers){                          //n4�p
      document.captureEvents(Event.MOUSEMOVE)
      document.captureEvents(Event.MOUSEUP)
    }
  }

  //--�C�x���g�L���v�`���[��~
  function stop_event__wrtCalendarLay(){
    document.onmousemove = null                   //n4,m1,n6,e4,e5,e6,o6�p
    document.onmouseup   = null                   //n4,m1,n6,e4,e5,e6,o6�p
    if(navigator.userAgent.indexOf('Gecko')!=-1)  //m1,n6�p
      document.onmousedown = null
    if(document.layers){                          //n4�p
      document.releaseEvents(Event.MOUSEMOVE)
      document.releaseEvents(Event.MOUSEUP)
    }
  }

  //--�u���E�U�̌�����擾
  function getBrowserLANG(){
    if(document.all)                  
      return navigator.browserLanguage      //e4,e5,e6,o6�p
    else if(document.layers) 
      return navigator.language             //n4�p
    else if(document.getElementById) 
      return navigator.language.substr(0,2) //n6,n7,m1�p
  }



  /*--/////////////�����܂�///////////////////////////////////////--*/

