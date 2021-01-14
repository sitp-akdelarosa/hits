//ヘルプ画面表示
//Input :int 1:実搬出　2:空搬入　3:空搬出　4:実搬入
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

//輸入コンテナ情報表示
//Input :フォームオブジェクト,詳細(1)・一覧(1以外)フラグ,新ウインドウの真偽
//Output:偽
//新規ウインドウを開く
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

//輸出コンテナ情報表示
//Input :フォームオブジェクト
//Output:
//新規ウインドウを開く
function BookInfo(target){
  target.action="./dmo930.asp"
  newWin = window.open("","ConInfo","left=30,top=10,status=yes,scrollbars=yes,resizable=yes,menubar=yes,width=1600,height=600");
  target.target="ConInfo";
  target.submit();
  target.target="_self";
}

//左の空白を削除する
//Input :ストリング
//Output:左の空白を削除したストリング
function LTrim(strTemp)
{
    var nLoop = 0;
    var strReturn = strTemp;
    while (nLoop < strTemp.length){
      if ((strReturn.substring(0, 1) == " ") || (strReturn.substring(0, 1) == "　"))
        strReturn = strTemp.substring(nLoop + 1, strTemp.length);
      else break;
      nLoop++;
    }
    return strReturn;
}

//文字列中に半角英数字と記号以外が無いかチェックする
//Input :ストリング
//Output:ない⇒真
//　　　:ある⇒偽
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
//文字列中に半角英数字以外が無いかチェックする
//Input :ストリング
//Output:ない⇒真
//　　　:ある⇒偽
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
//文字列中に半角英数字とその他の許可文字以外が無いかチェックする
//Input :ストリング
//Output:ない⇒真
//　　　:ある⇒偽
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
//文字列中に数字以外が無いかチェックする
//Input :ストリング
//Output:ない⇒真
//　　　:ある⇒偽
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
//文字列中に半角数字とその他の許可文字以外が無いかチェックする
//Input :ストリング
//Output:ない⇒真
//　　　:ある⇒偽
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

//閏年チェック
//INPUT：YYYY
//Output:閏年⇒真
//　　　:平年⇒偽
function isURU(year){
  if((year % 4 == 0 && year % 100 != 0) || year % 400 == 0)
    return true;
  else
    return false;
}

//閏年調整
//INPUT：YYYY(今年),MM(今月),document.form.select.Month,document.form.select.Day
function check_date(YYYY,MM,targetM,targetD){
  Month = targetM.selectedIndex;
  Dindex= targetD.selectedIndex;
  len   = targetD.length;
  if( Month < MM ){
    //選択された月が今月より小さい場合来年とみなす
    YYYY=Number(YYYY)+1;
  }
  if(Month==2){  //2月ならば閏年チェックを行う
    if(isURU(YYYY)){
      //日付のリストを29日に調整する
      for(i=len;i>30;i--)
        targetD.options[i-1]=null;
    } else {
      //日付のリストを28日に調整する
      for(i=len;i>29;i--)
        targetD.options[i-1]=null;
    }
  } else if(Month==4 || Month==6 || Month==9 || Month==11){
    //日付のリストを30日に調整する
    if(len<32){
      for(i=len-1;i<=30;i++)
        targetD.options[i]=new Option(i,i);
    } else {
        targetD.options[len-1]=null;
    }
  } else {
    //日付のリストを31日に調整する
    for(i=len-1;i<=31;i++)
      targetD.options[i]=new Option(i,i);
  }
  len=targetD.length
  if(Dindex>len-1)
    targetD.selectedIndex=len-1;
  else 
    targetD.selectedIndex=Dindex;
}

//指定した選択リストに[ ]、01～31の日付を入れindexで指定された日付をデフォルトにする
//INPUT：window.document.form.select,Number
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
//指定した選択リストに[ ]、01～12の月付を入れindexで指定された月付をデフォルトにする
//INPUT：window.document.form.select,Number 
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

//指定した選択リストに与えられた値入れindexで指定された値をデフォルトにする
//INPUT：window.document.form.select,Array,char
function setList(target,list,index){
  for(i=0;i<list.length;i++){
    target.options[i] = new Option(list[i],list[i]);
    if(list[i]==index)
        target.selectedIndex=i;
  }
}

//指定されたフォームの入力が禁止されている項目のバックグラウンドを変更する
//INPUT：window.document.form
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

//指定されたフォームの入力文字をすべて大文字に変更する
//INPUT：window.document.form
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
//ヘッダIDの制御
//INPUT：タイプ、window.document.form、対象会社コード、ログインユーザコード
//　　 ：タイプ：0→会社コードを変更　1→ヘッダコードを変更
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

//文字列のバイト数を計算する
//Input :ストリング
//Output:Array(バイト数,半角文字数,全角文字数)
function getByte(text)
{
  checkstr="ｱｲｳｴｵｶｷｸｹｺｻｷｽｾｿﾀﾁﾂﾃﾄｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｯｬｭｮﾞﾟ";
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

//文字列中に不正記号が無いかチェックする
//Input :ストリング
//Output:ない⇒真
//　　　:ある⇒偽
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
//日付の正当性チェック
//Input :NowYear(今年),NowMon(今月),targetMM,targetDD,targetHH
//Output:正当⇒真
//　　　:不当⇒偽
function CheckDate(NowYear,NowMon,targetMM,targetDD,targetHH){
  MM=targetMM.value;
  DD=targetDD.value;
  HH=targetHH.value;
  //Nullチェック
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
  
  //文字チェック
  if(!CheckSu(MM)){
     alert("半角数字以外の文字を入力しないでください");
     targetMM.focus();
     return false;
  }
  if(!CheckSu(DD)){
     alert("半角数字以外の文字を入力しないでください");
     targetDD.focus();
     return false;
  }
  if(!CheckSu(HH)){
     alert("半角数字以外の文字を入力しないでください");
     targetHH.focus();
     return false;
  }
  //期間チェック
  //月
  if(MM<1 || MM>12){
     alert("月は1～12の数字を入力してください");
     targetMM.focus();
     return false;
  }
  //日
  if( NowMon > MM ){
    //選択された月が今月より小さい場合来年とみなす
    NowYear=Number(NowYear)+1;
  }
  if(targetMM.value==2){  //2月ならば閏年チェックを行う
    if(isURU(NowYear)){
      //閏年
      MaxDay=29;
    } else {
      //平年
      MaxDay=28;
    }
  } else if(MM==4 || MM==6 || MM==9 || MM==11){
      MaxDay=30;
  } else {
      MaxDay=31;
  }
  if(DD<1 || DD>MaxDay){
     alert(MM+"月なので、日は1～"+ MaxDay +"の数字を入力してください");
     targetDD.focus();
     return false;
  }
  //時
  if(HH<0 || HH>23){
     alert("時は0～23の数字を入力してください");
     targetHH.focus();
     return false;
  }
  return true;
}
//ADD 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
//指定フォームの上からEnd-numまでreadOnlyにする
//Input :targetobj,num
function allsetreadOnly(target,num){
  len=target.elements.length;
  for(i=0;i<len-num;i++){
    target.elements[i].readOnly=true;
    if(document.dmi320F.elements[i].type == "checkbox" || document.dmi320F.elements[i].type == "select-one"){
    	target.elements[i].disabled=true;
    }
    if(document.dmi320F.elements[i].type == "button" && document.dmi320F.elements[i].value == "検索"){
    	target.elements[i].disabled=true;
    } 

  }
}

//ADD 20080131 START for Minutes By SITP G.Ariola
//日付の正当性チェック
//Input :NowYear(今年),NowMon(今月),targetMM,targetDD,targetHH,targetMN
//Output:正当⇒真
//　　　:不当⇒偽
function CheckDatewithMin(NowYear,NowMon,targetMM,targetDD,targetHH,targetMN){
  MM=targetMM.value;
  DD=targetDD.value;
  HH=targetHH.value;
  MN=targetMN.value;
  //Nullチェック
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
  
  //文字チェック
  if(!CheckSu(MM)){
     alert("半角数字以外の文字を入力しないでください");
     targetMM.focus();
     return false;
  }
  if(!CheckSu(DD)){
     alert("半角数字以外の文字を入力しないでください");
     targetDD.focus();
     return false;
  }
  if(!CheckSu(HH)){
     alert("半角数字以外の文字を入力しないでください");
     targetHH.focus();
     return false;
  }
  if(!CheckSu(MN)){
     alert("半角数字以外の文字を入力しないでください");
     targetMN.focus();
     return false;
  }
  //期間チェック
  //月
  if(MM<1 || MM>12){
     alert("月は1～12の数字を入力してください");
     targetMM.focus();
     return false;
  }
  //日
  if( NowMon > MM ){
    //選択された月が今月より小さい場合来年とみなす
    NowYear=Number(NowYear)+1;
  }
  if(targetMM.value==2){  //2月ならば閏年チェックを行う
    if(isURU(NowYear)){
      //閏年
      MaxDay=29;
    } else {
      //平年
      MaxDay=28;
    }
  } else if(MM==4 || MM==6 || MM==9 || MM==11){
      MaxDay=30;
  } else {
      MaxDay=31;
  }
  if(DD<1 || DD>MaxDay){
     alert(MM+"月なので、日は1～"+ MaxDay +"の数字を入力してください");
     targetDD.focus();
     return false;
  }
  //時
  if(HH<0 || HH>23){
     alert("時は0～23の数字を入力してください");
     targetHH.focus();
     return false;
  }
  //分
  if(MN<0 || MN>59){
     alert("分は0～59の数字を入力してください");
     targetHH.focus();
     return false;
  }
  return true;
}
