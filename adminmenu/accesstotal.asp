<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<%
'**********************************************
'  �y�v���O�����h�c�z�@: 
'  �y�v���O�������́z�@: 
'
'  �i�ύX�����j
'	2010/01/28	C.Pestano	��ʂɺ��Ēǉ�
'**********************************************
	
	Option Explicit
	Response.Expires = 0

	call CheckLoginH()
%>
<!--#include File="./Common/Common.inc"-->
<SCRIPT src="./Common/function.js" type=text/javascript></SCRIPT>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>���p�����\��</TITLE>
<link href="./Common/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

//**************************************************
// �T�v�@ : �����񒆂ɐ����ȊO���������`�F�b�N����
// �����@ : �X�g�����O
// �߂�l : true(�Ȃ�)/false(����)
//**************************************************
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

//**************************************************
// �T�v�@ : �[�N�`�F�b�N
// �����@ : YYYY
// �߂�l : true(�[�N)/false(���N)
//**************************************************
function isURU(year){
  if((year % 4 == 0 && year % 100 != 0) || year % 400 == 0)
    return true;
  else
    return false;
}
//**************************************************
// �T�v�@ : ���t�̐������`�F�b�N
// �����@ : targetYYYY,targetMM,targetDD,Mode(1:�N���� 2:�N��)
// �߂�l : true(����)/false(�s��)
//**************************************************
function CheckDate(targetYYYY,targetMM,targetDD,Mode){
  YYYY=targetYYYY.value;
  MM=targetMM.value;
  if (Mode==1){
    DD=targetDD.value;
  }else{
    DD=1;
  }
  //Null�`�F�b�N
  if(YYYY==null || YYYY==""){
    alert("�N�͕K�{���͍��ڂł��B");
    targetYYYY.focus();
    return false;
  }else if(MM==null || MM==""){
    alert("���͕K�{���͍��ڂł��B");
    targetMM.focus();
    return false;
  }else if(DD==null || DD==""){
    if (Mode==1){
       alert("���͕K�{���͍��ڂł��B");
       targetDD.focus();
       return false;
    }
  }
  
  //�����`�F�b�N
  if(!CheckSu(YYYY)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetYYYY.focus();
     return false;
  }else if(!CheckSu(MM)){
     alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
     targetMM.focus();
     return false;
  }
  if(!CheckSu(DD)){
    if (Mode==1){
        alert("���p�����ȊO�̕�������͂��Ȃ��ł�������");
        targetDD.focus();
        return false;
    }
  }
  //���ԃ`�F�b�N
  //��
  if(MM<1 || MM>12){
     alert("����1�`12�̐�������͂��Ă�������");
     targetMM.focus();
     return false;
  }
  //��
  if(targetMM.value==2){  //2���Ȃ�Ή[�N�`�F�b�N���s��
    if(isURU(YYYY)){
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
     alert(YYYY+"�N"+MM+"���Ȃ̂ŁA����1�`"+ MaxDay +"�̐�������͂��Ă�������");
     targetDD.focus();
     return false;
  }
  return true;
}

//���Ԍ����̃`�F�b�N
function CheckList(){
	var FromDate;
	var ToDate;
	//���̓`�F�b�N
	if(document.frm.rdoDetail[0].checked){
		//���t�̃`�F�b�NFrom
		if(CheckDate(document.frm.txtSYearFrom,document.frm.txtSMonthFrom,document.frm.txtSDayFrom,1)==true){
			FromDate=document.frm.txtSYearFrom.value;
			if(document.frm.txtSMonthFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSMonthFrom.value;
			}else{
				FromDate=FromDate+"/"+document.frm.txtSMonthFrom.value;
			}
			if(document.frm.txtSDayFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSDayFrom.value;	
			}else{
				FromDate=FromDate+"/"+document.frm.txtSDayFrom.value;
			}
		}else{
			return false;
		}
		//���t�̃`�F�b�NTo
		if(CheckDate(document.frm.txtSYearTo,document.frm.txtSMonthTo,document.frm.txtSDayTo,1)==true){
			ToDate=document.frm.txtSYearTo.value;
			if(document.frm.txtSMonthTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSMonthTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSMonthTo.value;
			}
			if(document.frm.txtSDayTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSDayTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSDayTo.value;
			}
		}else{
			return false;
		}
	}else{
		//���t�̃`�F�b�NFrom
		if(CheckDate(document.frm.txtSYearFrom,document.frm.txtSMonthFrom,"",2)==true){
			FromDate=document.frm.txtSYearFrom.value;
			if(document.frm.txtSMonthFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSMonthFrom.value;
			}else{
				FromDate=FromDate+"/"+document.frm.txtSMonthFrom.value;
			}
			FromDate=FromDate+"/01"
		}else{
			return false;
		}
		//���t�̃`�F�b�NTo
		if(CheckDate(document.frm.txtSYearTo,document.frm.txtSMonthTo,"",2)==true){
			ToDate=document.frm.txtSYearTo.value;
			if(document.frm.txtSMonthTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSMonthTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSMonthTo.value;
			}
			ToDate=ToDate+"/01"
		}else{
			return false;
		}
	}
	//���t�̑O��`�F�b�N
	if(FromDate > ToDate){
		alert("���t�͈̔͂Ɍ�肪����܂��B");
		document.frm.txtSYearFrom.focus();
		return false;
	}
return true;
}

//�݌v�\�̃`�F�b�N
function CheckTotal(){
	var FromDate;
	var ToDate;
	//���̓`�F�b�N
	//���t�̃`�F�b�NFrom
	if(CheckDate(document.frm.txtRYearFrom,document.frm.txtRMonthFrom,"",2)==true){
		FromDate=document.frm.txtRYearFrom.value;
		if(document.frm.txtRMonthFrom.value.length==1){
			FromDate=FromDate+"/0"+document.frm.txtRMonthFrom.value;
		}else{
			FromDate=FromDate+"/"+document.frm.txtRMonthFrom.value;
		}
	}else{
		return false;
	}
	//���t�̃`�F�b�NTo
	if(CheckDate(document.frm.txtRYearTo,document.frm.txtRMonthTo,"",2)==true){
		ToDate=document.frm.txtRYearTo.value;
		if(document.frm.txtRMonthTo.value.length==1){
			ToDate=ToDate+"/0"+document.frm.txtRMonthTo.value;
		}else{
			ToDate=ToDate+"/"+document.frm.txtRMonthTo.value;
		}
	}else{
		return false;
	}
	//���t�̑O��`�F�b�N
	if(FromDate > ToDate){
		alert("���t�͈̔͂Ɍ�肪����܂��B");
		document.frm.txtRYearFrom.focus();
		return false;
	}
return true;
}
// ���Ԍ�����������
function fListSearch(){
	var FromDate;
	var ToDate;
	var Mode;

	ret=CheckList();
	if (ret==true){	
		//From���t�쐬
		FromDate=document.frm.txtSYearFrom.value;
		if(document.frm.txtSMonthFrom.value.length==1){
			FromDate=FromDate+"/0"+document.frm.txtSMonthFrom.value;
		}else{
			FromDate=FromDate+"/"+document.frm.txtSMonthFrom.value;
		}
		//Mode�ݒ肪���ʂ̏ꍇ
		if(document.frm.rdoDetail[0].checked){
			if(document.frm.txtSDayFrom.value.length==1){
				FromDate=FromDate+"/0"+document.frm.txtSDayFrom.value;	
			}else{
				FromDate=FromDate+"/"+document.frm.txtSDayFrom.value;
			}
		}else{
			FromDate=FromDate+"/01"
		}
		//To���t�쐬
		ToDate=document.frm.txtSYearTo.value;
		if(document.frm.txtSMonthTo.value.length==1){
			ToDate=ToDate+"/0"+document.frm.txtSMonthTo.value;
		}else{
			ToDate=ToDate+"/"+document.frm.txtSMonthTo.value;
		}
		//Mode�ݒ肪���ʂ̏ꍇ
		if(document.frm.rdoDetail[0].checked){
			if(document.frm.txtSDayTo.value.length==1){
				ToDate=ToDate+"/0"+document.frm.txtSDayTo.value;
			}else{
				ToDate=ToDate+"/"+document.frm.txtSDayTo.value;
			}
		}else{
			ToDate=ToDate+"/01"
		}
		//Mode�ݒ�
		if(document.frm.rdoDetail[0].checked){
			Mode="D"
		}else{
			Mode="M"
		}
		document.frm.action="./logview.asp?fDate="+FromDate+"&tDate="+ToDate+"&Mode="+Mode;
		document.frm.submit();
		return true;
	}
	return ;

}

// �݌v�\��������
function fTotalSearch(){
	var FromDate;
	var ToDate;

	ret=CheckTotal();
	if (ret==true){	
		//From���t�쐬
		FromDate=document.frm.txtRYearFrom.value;
		if(document.frm.txtRMonthFrom.value.length==1){
			FromDate=FromDate+"/0"+document.frm.txtRMonthFrom.value;
		}else{
			FromDate=FromDate+"/"+document.frm.txtRMonthFrom.value;
		}
		//To���t�쐬
		ToDate=document.frm.txtRYearTo.value;
		if(document.frm.txtRMonthTo.value.length==1){
			ToDate=ToDate+"/0"+document.frm.txtRMonthTo.value;
		}else{
			ToDate=ToDate+"/"+document.frm.txtRMonthTo.value;
		}
		document.frm.action="./logListview.asp?fDate="+FromDate+"&tDate="+ToDate;
		document.frm.submit();
		return true;
	}
	return false;

}
</script>

</script>
</HEAD>
<body class="bckcolor" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="frm" action="accesstotal.asp" method="post" enctype="multipart/form-data">
<!-------------�������烍�O�C�����͉��--------------------------->
<table class="main2" align="center" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign=top>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<%
			DisplayHeader2("���p�����\��")
		  %>
		</table>
		<center>
		<table class="square" border="0" cellspacing="4" cellpadding="0" >
			<tr>
			<td>
			<table border="0" cellspacing="3" cellpadding="4">
				<tr>
				<td>
				<table width="100%" border="0" cellspacing="2" cellpadding="3">
					<tr> 
					<td colspan="3" align=left valign=center nowrap>�P�D���Ԍ���</td>
					</tr>
					<tr>
					<td width="50"></td>
					<td>
					<P>
					<INPUT name="txtSYearFrom" Type="Text" size="3" maxlength="4" style="WIDTH: 35px; HEIGHT: 20px" >&nbsp; �N &nbsp;
					<INPUT name="txtSMonthFrom" Type="Text" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp; ��&nbsp; 
					<INPUT name="txtSDayFrom" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp; ��&nbsp;&nbsp;����&nbsp; 
					<INPUT name="txtSYearTo" size="3" maxlength="4" style="WIDTH: 35px; HEIGHT: 20px">&nbsp; �N&nbsp;
					<INPUT name="txtSMonthTo" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp;&nbsp;&nbsp;��&nbsp;
					<INPUT name="txtSDayTo" size="1" maxlength="2" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp; �� �܂�
					</P>
					<P><INPUT type=radio name="rdoDetail" value="1" checked>&nbsp; ���ʖ���</P>
					<P><INPUT type=radio name="rdoDetail" value="2">&nbsp; ���ʖ���</P>
					<P>���ߋ��R�N���̂݌����\�ł��B<BR>
					����ʏ�ł͎w����Ԃ�TOTAL�����̂ݕ\������܂��B<BR>&nbsp;&nbsp;&nbsp; 
					CSV�o�͂ł͓��ʁA�܂��́A���ʂ̖��ׂ��\������܂��B<BR>
					<!-- 2010/01/28 Add-S C.Pestano -->
					�����v���ɂ͕����s���v�f�[�^�𔽉f�����Ă��܂��B<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�i2000�N11���`2009�N10���j
					<!-- 2010/01/28 Add-E C.Pestano -->
					</P> 
					</td>
					<td valign="top">&nbsp;</td>
					</tr>
					<tr> 
					<td colspan="3" align=middle valign=center nowrap>
					<P>
					<INPUT style="WIDTH: 112px; HEIGHT: 29px" type="button" size="37" value="���Ԍ���" name="btnSearch" Onclick="fListSearch();"></P>
					</td>
					</tr>
					<TR>
					<td>
					<P>�Q�D�݌v�\</P>
					</td>
					<tr> 
					<td width="50"></td>
					<td colspan="2" nowrap>
					<P>
					<INPUT name="txtRYearFrom" maxLength="4" size="3" style="WIDTH: 35px; HEIGHT: 20px"> �N&nbsp;
					<INPUT name="txtRMonthFrom" maxLength="2" size="1" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp;��&nbsp;&nbsp;&nbsp;����&nbsp; 
					<INPUT name="txtRYearTo" maxLength="4" size="3" style="WIDTH: 35px; HEIGHT: 20px">�N&nbsp; 
					<INPUT name="txtRMonthTo" maxLength="2" size="1" style="LEFT: 75px; WIDTH: 23px; TOP: 2px; HEIGHT: 21px">&nbsp;��&nbsp; �܂�
					</P>
					<P>���R�N�ȓ��͌��P�ʁA�S�N�ȏ�͔N�P�ʂŕ\������܂��B<BR>
					��HiTS��2000�N11������̃X�^�[�g�ł��B<BR>
					��2000�N11���`2009�N10���̊��Ԃ́A�����s���v�f�[�^�𔽉f�����Ă��܂��B<!-- 2010/01/28 Add C.Pestano -->
					</P> 
					</td>
					<td></td>
					</tr>
					<TR>
					<td colspan="3" align=middle valign=center nowrap>
					<INPUT style="WIDTH: 112px; HEIGHT: 29px" type="button" size="37" value="�݌v�\" name="btnRuikei" Onclick="fTotalSearch();">
					<P></P>
					</td>
					<td></td></TR>
					</table>
					<center>
					<A href="menu.asp">����</A>
					</center></td>
					</tr>
					</table>
				</td>
				</tr>
			</table></center>
			</td>
			</tr>
		<%
			DisplayFooter
		%>
		</table>
		</td>
	</Tr>
</form>
</body>
</HTML>
