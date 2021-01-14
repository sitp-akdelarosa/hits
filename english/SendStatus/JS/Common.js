//�w���v��ʕ\��
function GoHelp(Target){
	switch (Target) {
		case 1:
			Win = window.open('sst_help.asp', 'HelpWin', 'width=850,height=500,menubar=yes,scrollbars=yes');
			break;
	}
}

//�A���R���e�i���\��
//Input :�t�H�[���I�u�W�F�N�g,�ڍ�(1)�E�ꗗ(1�ȊO)�t���O,�V�E�C���h�E�̐^�U
//Output:�U
//�V�K�E�C���h�E���J��
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

//�����񒆂ɔ��p�p�����A���p�X�y�[�X�A-�A/�ȊO���������`�F�b�N����
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
		}else{
			//���t�̃��X�g��28���ɒ�������
			for(i=len;i>29;i--)
				targetD.options[i-1]=null;
		}
	}else if(Month==4 || Month==6 || Month==9 || Month==11){
		//���t�̃��X�g��30���ɒ�������
		if(len<32){
			for(i=len-1;i<=30;i++)
				targetD.options[i]=new Option(i,i);
		}else{
				targetD.options[len-1]=null;
		}
	}else{
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
		if(target.elements[i].disabled){
			target.elements[i].style.backgroundColor="#00FF00";
		}
	}
}

//�w�肳�ꂽ�t�H�[���̓��͕��������ׂđ啶���ɕύX����
//INPUT�Fwindow.document.form
function changeUpper(target){
	len=target.elements.length;
	for(i=0;i<len;i++){
		if(target.elements[i].type=="text"){
			tmp=target.elements[i].value
			target.elements[i].value=tmp.toUpperCase();
		}
	}
}

//�w�b�_ID�̐���
//INPUT�F�^�C�v�Awindow.document.form�A�Ώۉ�ЃR�[�h�A���O�C�����[�U�R�[�h
//     �F�^�C�v�F0����ЃR�[�h��ύX�@1���w�b�_�R�[�h��ύX
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

//������̃o�C�g�����v�Z����
//Input :�X�g�����O
//Output:Array(�o�C�g��,���p������,�S�p������)
function getByte(text){
	retA = new Array(0,0,0);
	for (i=0; i<text.length; i++){
		n = escape(text.charAt(i));
		if (n.length < 4) retA[1]++; else retA[2]++;
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
//
