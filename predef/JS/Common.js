//�w���v��ʕ\��
//Input :int 1:�����o�@2:������@3:����o�@4:������
function GoHelp(Target){
  switch (Target) {
    case 1:
      Win = window.open('Help50.ASP', 'HelpWin', 'width=850,height=500,left=0,top=0,menubar=yes,scrollbars=yes');
//      top.location.href = "Help50.ASP";
      break;
    case 2:
      Win = window.open('Help51.ASP', 'HelpWin', 'width=850,height=500,left=0,top=0,menubar=yes,scrollbars=yes');
//      top.location.href = "Help51.ASP";
      break;
    case 3:
      Win = window.open('Help52.ASP', 'HelpWin', 'width=850,height=500,left=0,top=0,menubar=yes,scrollbars=yes');
//      top.location.href = "Help52.ASP";
      break;
    case 4:
      Win = window.open('Help53.ASP', 'HelpWin', 'width=850,height=500,left=0,top=0,menubar=yes,scrollbars=yes');
//      top.location.href = "Help53.ASP";
      break;
  }
}

//�A���R���e�i���\��
//Input :�t�H�[���I�u�W�F�N�g,�ڍ�(1)�E�ꗗ(1�ȊO)�t���O,�V�E�C���h�E�̐^�U
//Output:�U
//�V�K�E�C���h�E���J��
function ConInfo(targetF,flag,newW){
  if(newW==0){
    newWin = window.open("","ConInfo","left=30,top=10status=yes,scrollbars=yes,resizable=yes,menubar=yes,width=800,height=600");
    targetF.target="ConInfo";
  }
  else{
    window.resizeTo(800,600);
    targetF.target="_self";
  }
  if(flag==1) targetF.action="./dmo910.asp";
  else        targetF.action="./dmo920.asp";
  targetF.submit();
  targetF.target="_self";
  return false;
}

//�A�o�R���e�i���\��
//Input :�t�H�[���I�u�W�F�N�g
//Output:
//�V�K�E�C���h�E���J��
function BookInfo(target){
  target.action="./dmo930.asp"
  newWin = window.open("","ConInfo","left=30,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes,width=1600,height=600");
  target.target="ConInfo";
  target.submit();
  target.target="_self";
}

//���̋󔒂��폜����
//Input :�X�g�����O
//Output:���̋󔒂��폜�����X�g�����O
function LTrim(strTemp)
{
    var nLoop = 0;
    var strReturn = strTemp;
    while (nLoop < strTemp.length){
      if ((strReturn.substring(0, 1) == " ") || (strReturn.substring(0, 1) == "�@"))
        strReturn = strTemp.substring(nLoop + 1, strTemp.length);
      else break;
      nLoop++;
    }
    return strReturn;
}

//�����񒆂ɔ��p�p�����ƋL���ȊO���������`�F�b�N����
//Input :�X�g�����O
//Output:�Ȃ��ː^
//�@�@�@:����ˋU
function CheckEisu(str){
  checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz /-";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      return false;
    }
  }
  return true;
}
//�����񒆂ɔ��p�p�����ȊO���������`�F�b�N����
//Input :�X�g�����O
//Output:�Ȃ��ː^
//�@�@�@:����ˋU
function CheckEisu2(str){
  checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      return false;
    }
  }
  return true;
}
//2017/05/08 H.Yoshikawa Add Start
//�����񒆂ɔ��p�p�����Ƃ��̑��̋������ȊO���������`�F�b�N����
//Input :�X�g�����O
//Output:�Ȃ��ː^
//�@�@�@:����ˋU
function CheckEisu3(str, kyoka){
  checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      if(kyoka.indexOf(c,0) < 0){
        return false;
      }
    }
  }
  return true;
}
//2017/05/08 H.Yoshikawa Add End
//�����񒆂ɐ����ȊO���������`�F�b�N����
//Input :�X�g�����O
//Output:�Ȃ��ː^
//�@�@�@:����ˋU
function CheckSu(str){
  checkstr="0123456789";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      return false;
    }
  }
  return true;
}
//2017/06/22 H.Yoshikawa Add Start
//�����񒆂ɔ��p�����Ƃ��̑��̋������ȊO���������`�F�b�N����
//Input :�X�g�����O
//Output:�Ȃ��ː^
//�@�@�@:����ˋU
function CheckSu2(str, kyoka){
  checkstr="0123456789";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) < 0){
      if(kyoka.indexOf(c,0) < 0){
        return false;
      }
    }
  }
  return true;
}
//2017/06/22 H.Yoshikawa Add End

//�[�N�`�F�b�N
//INPUT�FYYYY
//Output:�[�N�ː^
//�@�@�@:���N�ˋU
function isURU(year){
  if((year % 4 == 0 && year % 100 != 0) || year % 400 == 0)
    return true;
  else
    return false;
}

//�[�N����
//INPUT�FYYYY(���N),MM(����),document.form.select.Month,document.form.select.Day
function check_date(YYYY,MM,targetM,targetD){
  Month = targetM.selectedIndex;
  Dindex= targetD.selectedIndex;
  len   = targetD.length;
  if( Month < MM ){
    //�I�����ꂽ����������菬�����ꍇ���N�Ƃ݂Ȃ�
    YYYY=Number(YYYY)+1;
  }
  if(Month==2){  //2���Ȃ�Ή[�N�`�F�b�N���s��
    if(isURU(YYYY)){
      //���t�̃��X�g��29���ɒ�������
      for(i=len;i>30;i--)
        targetD.options[i-1]=null;
    } else {
      //���t�̃��X�g��28���ɒ�������
      for(i=len;i>29;i--)
        targetD.options[i-1]=null;
    }
  } else if(Month==4 || Month==6 || Month==9 || Month==11){
    //���t�̃��X�g��30���ɒ�������
    if(len<32){
      for(i=len-1;i<=30;i++)
        targetD.options[i]=new Option(i,i);
    } else {
        targetD.options[len-1]=null;
    }
  } else {
    //���t�̃��X�g��31���ɒ�������
    for(i=len-1;i<=31;i++)
      targetD.options[i]=new Option(i,i);
  }
  len=targetD.length
  if(Dindex>len-1)
    targetD.selectedIndex=len-1;
  else 
    targetD.selectedIndex=Dindex;
}

//�w�肵���I�����X�g��[ ]�A01�`31�̓��t�����index�Ŏw�肳�ꂽ���t���f�t�H���g�ɂ���
//INPUT�Fwindow.document.form.select,Number
function setDate(target,index){
  if(index == "") index = 0;
  target.options[0] = new Option(" ",0);
  for(i=1;i<32;i++){
    if(i<10)
      target.options[i] = new Option("0"+i,i);
    else
      target.options[i] = new Option(i,i);
  }
  target.selectedIndex=Number(index);
}
//�w�肵���I�����X�g��[ ]�A01�`12�̌��t�����index�Ŏw�肳�ꂽ���t���f�t�H���g�ɂ���
//INPUT�Fwindow.document.form.select,Number 
function setMonth(target,index){
  if(index == "") index = 0;
  target.options[0] = new Option(" ",0);
  for(i=1;i<13;i++){
    if(i<10)
      target.options[i] = new Option("0"+i,i);
    else
      target.options[i] = new Option(i,i);
  }
  target.selectedIndex=Number(index);
}
//  today= new Date();
//  dd = today.getDate();
//  mm = today.getMonth();

//�w�肵���I�����X�g�ɗ^����ꂽ�l����index�Ŏw�肳�ꂽ�l���f�t�H���g�ɂ���
//INPUT�Fwindow.document.form.select,Array,char
function setList(target,list,index){
  for(i=0;i<list.length;i++){
    target.options[i] = new Option(list[i],list[i]);
    if(list[i]==index)
        target.selectedIndex=i;
  }
}

//�w�肳�ꂽ�t�H�[���̓��͂��֎~����Ă��鍀�ڂ̃o�b�N�O���E���h��ύX����
//INPUT�Fwindow.document.form
function bgset(target){
  len=target.elements.length;
  for(i=0;i<len;i++){
    if(target.elements[i].readOnly){
       target.elements[i].style.border="1px inset #dddddd";
       target.elements[i].style.backgroundColor="#dddddd";
       target.elements[i].style.color="#000000";
    }
    //2016.08.25 H.Yoshikawa Add Start
    else{
      if(target.elements[i].type=="text"){
        target.elements[i].style.border="1px solid gray";
        target.elements[i].style.backgroundColor="#ffffff";
        target.elements[i].style.color="#000000";
      }
    }
    //2016.08.25 H.Yoshikawa Add End
  }
}

//�w�肳�ꂽ�t�H�[���̓��͕��������ׂđ啶���ɕύX����
//INPUT�Fwindow.document.form
function chengeUpper(target){
  len=target.elements.length;
  for(i=0;i<len;i++){
    if(target.elements[i].type=="text"){
      if(target.elements[i].name.toUpperCase().indexOf("MAIL") < 0 ){
       tmp=target.elements[i].value
       target.elements[i].value=tmp.toUpperCase();
      }
    }
  }
}
//CW-017 ADD START
//�w�b�_ID�̐���
//INPUT�F�^�C�v�Awindow.document.form�A�Ώۉ�ЃR�[�h�A���O�C�����[�U�R�[�h
//�@�@ �F�^�C�v�F0����ЃR�[�h��ύX�@1���w�b�_�R�[�h��ύX
function checkID(type,target,targetCOMPcd,COMPcd){
  flag=true;
  if(type==0){
    if(targetCOMPcd.value.length!=0 && targetCOMPcd.value.toUpperCase()!=COMPcd){
      target.HedId.value="";
      target.HedId.readOnly=true;
      target.HedId.style.backgroundColor="#dddddd";
      target.HedId.style.Color="#000000";
    } else {
      target.HedId.readOnly=false;
      target.HedId.style.backgroundColor="#ffffff";
    }
  } else {
    if(target.HedId.value.length!=0 && targetCOMPcd.value.toUpperCase()!=COMPcd && target.CMPcd1.value.length!=0){
      targetCOMPcd.value="";
      targetCOMPcd.readOnly=true;
      targetCOMPcd.style.backgroundColor="#dddddd";
      target.HedId.style.Color="#000000";
    } else {
      targetCOMPcd.readOnly=false;
      targetCOMPcd.style.backgroundColor="#ffffff";
    }
  }
}
//CW-017 ADD END
//C-002 ADD

//������̃o�C�g�����v�Z����
//Input :�X�g�����O
//Output:Array(�o�C�g��,���p������,�S�p������)
function getByte(text)
{
  checkstr="�������������������Ļ��������������������������������ܦݧ����������";
  retA = new Array(0,0,0);
  for (i=0; i<text.length; i++)
  {
    n = escape(text.charAt(i));
    if (n.length < 4){ retA[1]++; }
    else{
     if (checkstr.indexOf(text.charAt(i),0) >= 0){
       retA[1]++;
     }else{
       retA[2]++;
     }
    }
  }
  retA[0]=retA[1]+retA[2]*2;
  return retA;
}

//�����񒆂ɕs���L�����������`�F�b�N����
//Input :�X�g�����O
//Output:�Ȃ��ː^
//�@�@�@:����ˋU
function CheckKin(str){
  checkstr="\"\'\\\~,.#$%&|!@*+;:?";
  for (i=0; i<str.length; i++){
    c = str.charAt(i);
    if (checkstr.indexOf(c,0) >= 0){
      return false;
    }
  }
  return true;
}
//���t�̐������`�F�b�N
//Input :NowYear(���N),NowMon(����),targetMM,targetDD,targetHH
//Output:�����ː^
//�@�@�@:�s���ˋU
function CheckDate(NowYear,NowMon,targetMM,targetDD,targetHH){
  MM=targetMM.value;
  DD=targetDD.value;
  HH=targetHH.value;
  //Null�`�F�b�N
  if(MM==null || MM==""){
    targetDD.value="";
    targetHH.value="";
    return true;
  }else if(DD==null || DD==""){
    targetMM.value="";
    targetHH.value="";
    return true;
  }else if(HH==null || HH==""){
    HH=0;
  }
  
  //�����`�F�b�N
  if(!CheckSu(MM)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetMM.focus();
     return false;
  }
  if(!CheckSu(DD)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetDD.focus();
     return false;
  }
  if(!CheckSu(HH)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetHH.focus();
     return false;
  }
  //���ԃ`�F�b�N
  //��
  if(MM<1 || MM>12){
     alert("����1�`12�̐�������͂��Ă�������");
     targetMM.focus();
     return false;
  }
  //��
  if( NowMon > MM ){
    //�I�����ꂽ����������菬�����ꍇ���N�Ƃ݂Ȃ�
    NowYear=Number(NowYear)+1;
  }
  if(targetMM.value==2){  //2���Ȃ�Ή[�N�`�F�b�N���s��
    if(isURU(NowYear)){
      //�[�N
      MaxDay=29;
    } else {
      //���N
      MaxDay=28;
    }
  } else if(MM==4 || MM==6 || MM==9 || MM==11){
      MaxDay=30;
  } else {
      MaxDay=31;
  }
  if(DD<1 || DD>MaxDay){
     alert(MM+"���Ȃ̂ŁA����1�`"+ MaxDay +"�̐�������͂��Ă�������");
     targetDD.focus();
     return false;
  }
  //��
  if(HH<0 || HH>23){
     alert("����0�`23�̐�������͂��Ă�������");
     targetHH.focus();
     return false;
  }
  return true;
}
//ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
//�w��t�H�[���̏ォ��End-num�܂�readOnly�ɂ���
//Input :targetobj,num
function allsetreadOnly(target,num){
  len=target.elements.length;
  for(i=0;i<len-num;i++){
    target.elements[i].readOnly=true;
    if(document.dmi320F.elements[i].type == "checkbox" || document.dmi320F.elements[i].type == "select-one"){
    	target.elements[i].disabled=true;
    }
    if(document.dmi320F.elements[i].type == "button" && document.dmi320F.elements[i].value == "����"){
    	target.elements[i].disabled=true;
    } 

  }
}

//ADD 20080131 START for Minutes By SITP G.Ariola
//���t�̐������`�F�b�N
//Input :NowYear(���N),NowMon(����),targetMM,targetDD,targetHH,targetMN
//Output:�����ː^
//�@�@�@:�s���ˋU
function CheckDatewithMin(NowYear,NowMon,targetMM,targetDD,targetHH,targetMN){
  MM=targetMM.value;
  DD=targetDD.value;
  HH=targetHH.value;
  MN=targetMN.value;
  //Null�`�F�b�N
  if(MM==null || MM==""){
    targetDD.value="";
    targetHH.value="";
    return true;
  }else if(DD==null || DD==""){
    targetMM.value="";
    targetHH.value="";
    return true;
  }else if(HH==null || HH==""){
    HH=0;
  }else if(MN==null || MN==""){
    MN=0;
  }
  
  //�����`�F�b�N
  if(!CheckSu(MM)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetMM.focus();
     return false;
  }
  if(!CheckSu(DD)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetDD.focus();
     return false;
  }
  if(!CheckSu(HH)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetHH.focus();
     return false;
  }
  if(!CheckSu(MN)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetMN.focus();
     return false;
  }
  //���ԃ`�F�b�N
  //��
  if(MM<1 || MM>12){
     alert("����1�`12�̐�������͂��Ă�������");
     targetMM.focus();
     return false;
  }
  //��
  if( NowMon > MM ){
    //�I�����ꂽ����������菬�����ꍇ���N�Ƃ݂Ȃ�
    NowYear=Number(NowYear)+1;
  }
  if(targetMM.value==2){  //2���Ȃ�Ή[�N�`�F�b�N���s��
    if(isURU(NowYear)){
      //�[�N
      MaxDay=29;
    } else {
      //���N
      MaxDay=28;
    }
  } else if(MM==4 || MM==6 || MM==9 || MM==11){
      MaxDay=30;
  } else {
      MaxDay=31;
  }
  if(DD<1 || DD>MaxDay){
     alert(MM+"���Ȃ̂ŁA����1�`"+ MaxDay +"�̐�������͂��Ă�������");
     targetDD.focus();
     return false;
  }
  //��
  if(HH<0 || HH>23){
     alert("����0�`23�̐�������͂��Ă�������");
     targetHH.focus();
     return false;
  }
  //��
  if(MN<0 || MN>59){
     alert("����0�`59�̐�������͂��Ă�������");
     targetHH.focus();
     return false;
  }
  return true;
}
