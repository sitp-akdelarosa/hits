//ヘルプ画面表示
function GoHelp(Target){
	switch (Target) {
		case 1:
			Win = window.open('sst_help.asp', 'HelpWin', 'width=850,height=500,menubar=yes,scrollbars=yes');
			break;
	}
}

//輸入コンテナ情報表示
//Input :フォームオブジェクト,詳細(1)・一覧(1以外)フラグ,新ウインドウの真偽
//Output:偽
//新規ウインドウを開く
function ConInfo(targetF,flag,newW){
	if(newW==0){
		newWin = window.open("","ConInfo","status=yes,scrollbars=yes,resizable=yes,menubar=yes,width=800,height=600");
		targetF.target="ConInfo";
	}else{
		window.resizeTo(800,600);
		targetF.target="_self";
	}
	if(flag==1) targetF.action="./dmo910.asp";
	else        targetF.action="./dmo920.asp";
	targetF.CONnum.disabled = false;
	targetF.submit();
	targetF.CONnum.disabled = true;
	targetF.target="_self";
	return false;
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

//文字列中に半角英数字、半角スペース、-、/以外が無いかチェックする
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
function CheckEisuji(str){
	checkstr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
	for (i=0; i<str.length; i++){
		c = str.charAt(i);
		if (checkstr.indexOf(c,0) < 0){
			return false;
		}
	}
	return true;
}

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
		}else{
			//日付のリストを28日に調整する
			for(i=len;i>29;i--)
				targetD.options[i-1]=null;
		}
	}else if(Month==4 || Month==6 || Month==9 || Month==11){
		//日付のリストを30日に調整する
		if(len<32){
			for(i=len-1;i<=30;i++)
				targetD.options[i]=new Option(i,i);
		}else{
				targetD.options[len-1]=null;
		}
	}else{
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
		if(target.elements[i].disabled){
			target.elements[i].style.backgroundColor="#00FF00";
		}
	}
}

//指定されたフォームの入力文字をすべて大文字に変更する
//INPUT：window.document.form
function changeUpper(target){
	len=target.elements.length;
	for(i=0;i<len;i++){
		if(target.elements[i].type=="text"){
			tmp=target.elements[i].value
			target.elements[i].value=tmp.toUpperCase();
		}
	}
}

//ヘッダIDの制御
//INPUT：タイプ、window.document.form、対象会社コード、ログインユーザコード
//     ：タイプ：0→会社コードを変更　1→ヘッダコードを変更
function checkID(type,target,targetCOMPcd,COMPcd){
	flag=true;
	if(type==0){
		if(targetCOMPcd.value.length!=0 && targetCOMPcd.value.toUpperCase()!=COMPcd){
			target.HedId.value="";
			target.HedId.disabled=true;
			target.HedId.style.backgroundColor="#00ff00";
		} else {
			target.HedId.disabled=false;
			target.HedId.style.backgroundColor="#ffffff";
		}
	} else {
		if(target.HedId.value.length!=0 && targetCOMPcd.value.toUpperCase()!=COMPcd && target.CMPcd1.value.length!=0){
			targetCOMPcd.value="";
			targetCOMPcd.disabled=true;
			targetCOMPcd.style.backgroundColor="#00ff00";
		} else {
			targetCOMPcd.disabled=false;
			targetCOMPcd.style.backgroundColor="#ffffff";
		}
	}
}

//文字列のバイト数を計算する
//Input :ストリング
//Output:Array(バイト数,半角文字数,全角文字数)
function getByte(text){
	retA = new Array(0,0,0);
	for (i=0; i<text.length; i++){
		n = escape(text.charAt(i));
		if (n.length < 4) retA[1]++; else retA[2]++;
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
//
