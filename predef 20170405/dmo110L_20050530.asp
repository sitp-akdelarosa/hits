<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo110L.asp				_/
'_/	Function	:��������ꗗ��ʃ��X�g�o��		_/
'_/	Date		:2003/05/27				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-001 2003/07/29	CSV�o�͑Ή�	_/
'_/			:C-002 2003/07/29	���l���Ή�	_/
'_/			:C-004 2003/08/22	�\�������`	_/
'_/			:3th   2004/01/31	3���Ή�	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
'CW-055  Session.Contents.Remove("DateP")
'CW-055  Session.Contents.Remove("NumP")

'���[�U�f�[�^����
  dim USER, COMPcd
  USER   = UCase(Session.Contents("userid"))
  COMPcd = Session.Contents("COMPcd")
'INI�t�@�C�����ݒ�l���擾
  dim tmp(2),calcDate1
  getIni tmp
  calcDate1 = DateAdd("d", "-"&tmp(1), Date)

'�f�[�^�擾
  dim Num, DtTbl,i,j,SortFlag,SortKye

  If Request("SortFlag") = "" Then
    SortFlag = 0
  Else
    SortFlag = Request("SortFlag")
  End If

  '�\�[�g�P�[�X
  dim strWrer,ErrerM
  Select Case SortFlag
      Case "0" '�����\��:�����\������ɕ\��
        WriteLogH "b201", "��������O���ꗗ","01",""
        strWrer="AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,0
      Case "1" '�w���悪���Ɖ�̃R���e�i�ꗗ
        WriteLogH "b201", "��������O���ꗗ","03",""
        strWrer="AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,1
'3th          j=1
'3th          DtTbl(0)(8)=0
'3th          For i=1 To Num
'3th            If DtTbl(i)(6) = "��" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(j)(7)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "7" '�ۗ�
        WriteLogH "b201", "��������O���ꗗ","07",""
        strWrer="AND (DateDiff(day,ITC.WorkCompleteDate,'"&calcDate1&"')<=0 Or ITC.WorkCompleteDate IS Null) "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,2
'3th          j=1
'3th          DtTbl(0)(8)=0
'3th          For i=1 To Num
'3th            If DtTbl(i)(6) = "No" Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(j)(7)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "2" '�����������������ׂĕ\��
        WriteLogH "b201", "��������O���ꗗ","02",""
        strWrer="AND ITC.WorkCompleteDate IS Null "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,0
      Case "3" '�S���\��
        WriteLogH "b201", "��������O���ꗗ","04",""
        strWrer=" "
'3th          getData DtTbl,strWrer
          getData DtTbl,strWrer,0
      Case "4" '�R���e�i�ԍ��Ō���
          SortKye=Request("SortKye")
          WriteLogH "b201", "��������O���ꗗ","11",SortKye
'CW-055 Chenge Start
'          If Session.Contents("ConNum") = "" Then
'            jampErrerP "0","b201","11","������F�ꗗ����(�R���e�i�ԍ�)","001",""
'          Else
'            DtTbl=Session("DateT")
'            Num  =Session.Contents("ConNum")
'          End If
'3th chage          Get_Data Num,DtTbl
          strWrer = "AND ITC.ContNo LIKE '%" & SortKye & "'"
          getData DtTbl,strWrer,0
'CW-055 Chenge End
'3th          j=1
'3th          DtTbl(0)(8)=0
'3th          For i=1 To Num
'3th            If Right(DtTbl(i)(3),Len(SortKye))= SortKye Then
'3th              DtTbl(j)=DtTbl(i)
'3th              DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(j)(7)
'3th              j=j+1
'3th            End If
'3th          Next
'3th          Num=j-1
      Case "8" '�Ɖ��
          WriteLogH "b207", "��������O���Ɖ��","01",SortKye
'CW-055 Chenge Start
'          If Session.Contents("ConNum") = "" Then
'            jampErrerP "0","b207","01","������F�ꗗ�Ɖ��","001",""
'          Else
'            DtTbl=Session("DateT")
'            Num  =Session.Contents("ConNum")
'          End If
          Get_Data Num,DtTbl
'CW-055 Chenge End
        '�G���[�g���b�v�J�n
          on error resume next
        'DB�ڑ�
          dim ObjConn, ObjRS, StrSQL
          ConnDBH ObjConn, ObjRS
          For i=1 To Num
'CW-002            If DtTbl(i)(7) <> 0 Then
            If DtTbl(i)(7) <> 0 AND DtTbl(i)(4)="" AND DtTbl(i)(8)="��" Then
              StrSQL = "UPDATE hITReference SET UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01'," &_
                       "UpdtTmnl='"& USER &"', TruckerFlag"& DtTbl(i)(7) &"=1 "&_
                       "WHERE WkContrlNo IN (select WkContrlNo From hITCommonInfo "&_
                       "WHERE ContNo='"& DtTbl(i)(3) &"' AND WkType='2'  AND Process='R')"
              ObjConn.Execute(StrSQL)
              if err <> 0 then
                Set ObjRS = Nothing
                jampErrerPDB ObjConn,"0","b207","01","������F�ꗗ�Ɖ��","103","SQL:<BR>"&strSQL
              end if
            End If
          Next
        'DB�ڑ�����
          DisConnDBH ObjConn, ObjRS
        '�G���[�g���b�v����
          on error goto 0
          Response.Redirect "./dmo110L.asp"
  End Select
'CW-055  Session.Contents.Remove("DateT")
'CW-055  Session("DateT")=DtTbl
'CW-055  Session.Contents("ConNum")=Num
'CW-055  If Num=0 Then
'CW-055    Session.Contents("NullFlag")=0
'CW-055  Else
'CW-055    Session.Contents("NullFlag")=1
'CW-055  End If

'�f�[�^�擾�֐�
'3thFunction getData(DtTbl,strWrer)
Function getData(DtTbl,strWrer,DelType)
  ReDim DtTbl(1)
'C-002  DtTbl(0)=Array("���͓�","�����\���","�w����","�R���e�i�ԍ�","��������","�w����","�w����<BR>��","�Ɖ��Frag","�w�����ւ̉�","�D��","�D��","�T�C�Y","�ԋp��","�f�B�e���V����<BR>�t���[�^�C��")
'20030911  DtTbl(0)=Array("���͓�","�����\���","�w����","�R���e�i�ԍ�","��������","�w����","�w����<BR>��","�Ɖ��Frag","�w�����ւ̉�","�D��","�D��","�T�C�Y","�ԋp��","�f�B�e���V����<BR>�t���[�^�C��","���l")
  DtTbl(0)=Array("���͓�","�����\���","�w����","�R���e�i�ԍ�","��������","�w����","�w����<BR>��","�Ɖ��Frag","�w�����ւ̉�","�D��","�D��","�T�C�Y","�ԋp��","�f�B�e���V����<BR>�t���[�^�C��","���l","��ƊǗ��ԍ�")
  DtTbl(0)(8) =0
'3th Add Start
  Dim DelStr,DelTarget
  DelStr=Array("","��","No")
  DelTarget=Array(0,6,6)
'3th Add End

  '�G���[�g���b�v�J�n
    on error resume next
  'DB�ڑ�
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS

  '�Ώی����擾
    StrSQL = "SELECT count(WkContrlNo) AS CNUM FROM hITCommonInfo ITC "&_
             "WHERE WkType='2' AND (RegisterCode='"& USER &"' "&_
             "OR TruckerSubCode1='"& COMPcd &"' OR TruckerSubCode2='"&_
              COMPcd &"' OR TruckerSubCode3='"& COMPcd &"' OR TruckerSubCode4='"& COMPcd &"') AND Process='R' " &_
              strWrer 
   ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "2","b201","01","������F�ꗗ","101","SQL:<BR>"&StrSQL
      Exit Function
    end if
    Num = ObjRS("CNUM")
    ObjRS.close
    ReDim Preserve DtTbl(Num)

  '�f�[�^�擾
    StrSQL = "SELECT ITC.InputDate, ITC.WorkDate, ITC.RegisterCode, ITC.TruckerSubCode1, ITC.TruckerSubCode2, "&_
             "ITC.TruckerSubCode3, ITC.TruckerSubCode4, ITC.ContNo, ITC.WorkCompleteDate, "&_
             "ITC.WkContrlNo, ITC.Comment1, ITR.TruckerFlag1, "&_
             "ITR.TruckerFlag2, ITR.TruckerFlag3, ITR.TruckerFlag4, mV.ShipLine, mV.FullName, CNT.ContSize, "&_
             "INC.ReturnPlace, INC.DetentionFreeTime, mU.HeadCompanyCode, mU.UserType "&_
             "FROM ((((hITCommonInfo AS ITC INNER JOIN hITReference AS ITR ON ITC.WkContrlNo = ITR.WkContrlNo) "&_
             "INNER JOIN Container AS CNT ON ITC.ContNo = CNT.ContNo) "&_
             "LEFT JOIN mVessel AS mV ON CNT.VslCode = mV.VslCode) "&_
             "INNER JOIN ImportCont AS INC ON (CNT.ContNo = INC.ContNo) AND (CNT.VoyCtrl = INC.VoyCtrl) "&_
             "AND (CNT.VslCode = INC.VslCode))"&_
             "INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
             "WHERE WkType='2' AND (RegisterCode='"& USER &"' "&_
             "OR TruckerSubCode1='"& COMPcd &"' OR TruckerSubCode2='"&_
              COMPcd &"' OR TruckerSubCode3='"& COMPcd &"' OR TruckerSubCode4='"& COMPcd &"') AND Process='R' " &_
             strWrer &_
             "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC"
'CW-051 ADD This Line "mU.HeadCompanyCode, mU.UserType "&_
'CW-051 ADD This Line "INNER JOIN mUsers AS mU ON ITC.RegisterCode = mU.UserCode "&_
'C-004 Chenge This Line "ORDER BY isnull(ITC.WorkDate,DATEADD(Year,100,getdate())),ITC.InputDate ASC ASC"
'20030911 ADD This Item "ITC.WkContrlNo, "
'C-002 ADD 
    ObjRS.Open StrSQL, ObjConn
    if err <> 0 then
      DisConnDBH ObjConn, ObjRS	'DB�ؒf
      jampErrerP "2","b201","01","������F�ꗗ","102","SQL:<BR>"&StrSQL
      Exit Function
    end if
    i=1
    Do Until ObjRS.EOF
'C-002      DtTbl(i)=Array("","","","","","","","","","","","","","")
'20030911      DtTbl(i)=Array("","","","","","","","","","","","","","","")
      DtTbl(i)=Array("","","","","","","","","","","","","","","","")
      DtTbl(i)(0)=Mid(ObjRS("InPutDate"),3,8)
      DtTbl(i)(1)=Mid(ObjRS("WorkDate"),3,8)
      DtTbl(i)(3)=Trim(ObjRS("ContNo"))
      DtTbl(i)(4)=Trim(Mid(ObjRS("WorkCompleteDate"),3,8))
      DtTbl(i)(9)=Trim(ObjRS("ShipLine"))
'C-001      DtTbl(i)(10)=Left(ObjRS("FullName"),12)
      DtTbl(i)(10)=Trim(ObjRS("FullName"))
      DtTbl(i)(11)=Trim(ObjRS("ContSize"))
      DtTbl(i)(12)=Trim(ObjRS("ReturnPlace"))
      DtTbl(i)(13)=Trim(ObjRS("DetentionFreeTime"))
      DtTbl(i)(14)=Trim(ObjRS("Comment1"))		'C-002
      DtTbl(i)(15)=Trim(ObjRS("WkContrlNo"))		'20030911
    '�w����Ɖ�ς݃t���O
      IF Trim(ObjRS("TruckerSubCode4")) = COMPcd Then
        DtTbl(i)(2)=Trim(ObjRS("TruckerSubCode3"))
        DtTbl(i)(5)=Null
        DtTbl(i)(7)=4
        DtTbl(i)(6)=Null
        DtTbl(i)(8)=ObjRS("TruckerFlag4")
      ElseIF Trim(ObjRS("TruckerSubCode3")) = COMPcd Then
        DtTbl(i)(2)=Trim(ObjRS("TruckerSubCode2"))
        DtTbl(i)(5)=Trim(ObjRS("TruckerSubCode4"))
        DtTbl(i)(7)=3
        DtTbl(i)(6)=ObjRS("TruckerFlag4")
        DtTbl(i)(8)=ObjRS("TruckerFlag3")
      ElseIF Trim(ObjRS("TruckerSubCode2")) = COMPcd Then
        DtTbl(i)(2)=Trim(ObjRS("TruckerSubCode1"))
        DtTbl(i)(5)=Trim(ObjRS("TruckerSubCode3"))
        DtTbl(i)(7)=2
        DtTbl(i)(6)=ObjRS("TruckerFlag3")
        DtTbl(i)(8)=ObjRS("TruckerFlag2")
      ELSEIf Trim(ObjRS("TruckerSubCode1")) = COMPcd Then
        If ObjRS("UserType") = "5" Then			'CW-051
          DtTbl(i)(2)=Trim(ObjRS("HeadCompanyCode"))	'CW-051
        Else						'CW-051
          DtTbl(i)(2)=Trim(ObjRS("RegisterCode"))
        End If						'CW-051
        DtTbl(i)(5)=Trim(ObjRS("TruckerSubCode2"))
        DtTbl(i)(7)=1
        DtTbl(i)(6)=ObjRS("TruckerFlag2")
        DtTbl(i)(8)=ObjRS("TruckerFlag1")
      Else
        If ObjRS("UserType") = "5" Then			'CW-051
          DtTbl(i)(2)=Trim(ObjRS("HeadCompanyCode"))	'CW-051
        Else						'CW-051
          DtTbl(i)(2)=Trim(ObjRS("RegisterCode"))
        End If						'CW-051
        DtTbl(i)(5)=Trim(ObjRS("TruckerSubCode1"))
        DtTbl(i)(7)=0
        DtTbl(i)(6)=ObjRS("TruckerFlag1")
        DtTbl(i)(8)=Null
      End If
      If IsNull(DtTbl(i)(5)) Then
        DtTbl(i)(6) ="�@"
      ElseIf DtTbl(i)(6) = 0 Then
        DtTbl(i)(6) ="��"
      ElseIf DtTbl(i)(6) = 1 Then
        DtTbl(i)(6) ="Yes"
      Else
        DtTbl(i)(6) ="No"
      End If
      If DtTbl(i)(8)=0 Then
        DtTbl(i)(8) ="��"
      ElseIf DtTbl(i)(8) = 1 Then
        DtTbl(i)(8) ="Yes"
      ElseIf DtTbl(i)(8) = 2 Then
        DtTbl(i)(8) ="No"
      Else
        DtTbl(i)(8) ="�@"
      End If
'3th Add Start
      If DelType=0 OR DtTbl(i)(DelTarget(DelType)) = DelStr(DelType) Then
        DtTbl(0)(8) = DtTbl(0)(8) + DtTbl(i)(7)
        i=i+1
      Else
        Num=Num-1
      End If
'      i=i+1
'3th Add End
      ObjRS.MoveNext
    Loop
    ObjRS.close
    If i-1<Num Then
      ErrerM = "<DIV class=alert>�o�^�f�[�^�̂���"& Num-i+1 &"���ɂ��Ċ֘A�f�[�^�擾���s�̂���"&_
               "�\������Ă��܂���B<BR>�V�X�e���Ǘ��҂ɖ₢���킹�Ă��������B</DIV><P>"
      Num=i-1
    End If
  'DB�ڑ�����
    DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
    on error goto 0
End Function

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��������ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
//�f�[�^�������ꍇ�̕\������
function vew(){
}
//�X�V
function GoRenew(conNo,wkconNo){
  Fname=document.dmo110F;
  Fname.CONnum.value=conNo;
  Fname.WkconNo.value=wkconNo;
  Fname.action="./dmi115.asp";
  newWin = window.open("", "ReEntry", "status=yes,width=500,height=500,left=10,top=10,resizable=yes");
  Fname.target="ReEntry";
  Fname.submit();
}
//����
function SerchC(SortFlag,Kye){
  Fname=document.dmo110F;
  Fname.SortFlag.value=SortFlag;
  Fname.SortKye.value=Kye;
  Fname.target="_self";
  Fname.action="./dmo110L.asp";
  Fname.submit();
}
//�Ɖ��
function GoSyokaizumi(){
  target=document.dmo110F;
  if(target.DataNum.value>0){
    flag = confirm('���񓚂̉񓚂��uYes�v�ɂ��܂����H');
    if(flag==true){
      len=target.elements.length;
      for(i=0;i<len;i++){
        target.elements[i].disabled=false;
      }
      target.SortFlag.value=8;
      target.target="_self";
      target.action="./dmo110L.asp";
      target.submit();
    }
  }
}
//CSV		ADD C-001
function GoCSV(){
  target=document.dmo110F;
  len=target.elements.length;
  for(i=0;i<len;i++){
    target.elements[i].disabled=false;
  }
  target.target="Bottom";
  target.action="./dmo180.asp";
  target.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="vew()" >
<!-------------��������ꗗ���List--------------------------->
<%=ErrerM%>
<Form name="dmo110F" method="POST">
<TABLE border=1 cellPadding=3 cellSpacing=0 cols="<%=Num+1%>">
<%If Num<>0 Then%> 
  <% If DtTbl(0)(8)=0 Then %>
  <TR class=bga>
    <TH nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%></TH>
    <TH nowrap><%=DtTbl(0)(3)%></TH><TH nowrap><%=DtTbl(0)(9)%></TH><TH nowrap><%=DtTbl(0)(10)%></TH>
    <TH nowrap><%=DtTbl(0)(11)%></TH><TH nowrap><%=DtTbl(0)(12)%></TH><TH nowrap><%=DtTbl(0)(13)%></TH>
    <TH nowrap><%=DtTbl(0)(5)%></TH><TH nowrap><%=DtTbl(0)(6)%></TH><TH nowrap><%=DtTbl(0)(14)%>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>'>
    </TH>
  </TR>
    <% For j=1 to Num %>
  <TR class=bgw>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%></TD>
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(3)%>','<%=DtTbl(j)(15)%>');"><%=DtTbl(j)(3)%></A></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(10),12)%><BR></TD><TD nowrap><%=DtTbl(j)(11)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(12)%><BR></TD><TD nowrap><%=DtTbl(j)(13)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(5)%><BR></TD><TD nowrap><%=DtTbl(j)(6)%></TD><TD nowrap><%=Left(DtTbl(j)(14),10)%><BR>
    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>'>
    </TD>
  </TR>
    <% Next %>
  <% Else %>
  <TR class=bga>
    <TH nowrap><%=DtTbl(0)(1)%></TH><TH nowrap><%=DtTbl(0)(2)%></TH>
    <TH nowrap>�w����<BR>�ւ̉�</TH>
    <TH nowrap><%=DtTbl(0)(3)%></TH><TH nowrap><%=DtTbl(0)(9)%></TH><TH nowrap><%=DtTbl(0)(10)%></TH>
    <TH nowrap><%=DtTbl(0)(11)%></TH><TH nowrap><%=DtTbl(0)(12)%></TH><TH nowrap><%=DtTbl(0)(13)%></TH>
    </TH><TH nowrap><%=DtTbl(0)(5)%></TH><TH nowrap><%=DtTbl(0)(6)%></TH><TH nowrap><%=DtTbl(0)(14)%>
    <INPUT type=hidden name='Datatbl0' disabled value='<%=DtTbl(0)(0)%>,<%=DtTbl(0)(1)%>,<%=DtTbl(0)(2)%>,<%=DtTbl(0)(3)%>,<%=DtTbl(0)(4)%>,<%=DtTbl(0)(5)%>,<%=DtTbl(0)(6)%>,<%=DtTbl(0)(7)%>,<%=DtTbl(0)(8)%>,<%=DtTbl(0)(9)%>,<%=DtTbl(0)(10)%>,<%=DtTbl(0)(11)%>,<%=DtTbl(0)(12)%>,<%=DtTbl(0)(13)%>,<%=DtTbl(0)(14)%>'>
    </TH>
  </TR>
    <% For j=1 to Num %>
  <TR class=bgw>
    <TD nowrap><%=DtTbl(j)(1)%><BR></TD><TD nowrap><%=DtTbl(j)(2)%></TD>
    <TD nowrap><%=DtTbl(j)(8)%></TD> 
    <TD nowrap><A HREF="JavaScript:GoRenew('<%=DtTbl(j)(3)%>','<%=DtTbl(j)(15)%>');"><%=DtTbl(j)(3)%></A></TD>
    <TD nowrap><%=DtTbl(j)(9)%><BR></TD><TD nowrap><%=Left(DtTbl(j)(10),12)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(11)%><BR></TD><TD nowrap><%=DtTbl(j)(12)%><BR></TD><TD nowrap><%=DtTbl(j)(13)%><BR></TD>
    <TD nowrap><%=DtTbl(j)(5)%><BR></TD><TD nowrap><%=DtTbl(j)(6)%></TD><TD nowrap><%=Left(DtTbl(j)(14),10)%><BR>
    <INPUT type=hidden name='Datatbl<%=j%>' disabled value='<%=DtTbl(j)(0)%>,<%=DtTbl(j)(1)%>,<%=DtTbl(j)(2)%>,<%=DtTbl(j)(3)%>,<%=DtTbl(j)(4)%>,<%=DtTbl(j)(5)%>,<%=DtTbl(j)(6)%>,<%=DtTbl(j)(7)%>,<%=DtTbl(j)(8)%>,<%=DtTbl(j)(9)%>,<%=DtTbl(j)(10)%>,<%=DtTbl(j)(11)%>,<%=DtTbl(j)(12)%>,<%=DtTbl(j)(13)%>,<%=DtTbl(j)(14)%>'>
    </TD>
  </TR>
    <% Next %>
  <% End If %>
<% Else %>
  <TR class=bgw><TD nowrap>��ƈČ��͂���܂���</TD></TR>
<% End If %>
</TABLE>
<%'3th del Set_Data Num,DtTbl %>
  <INPUT type=hidden name=DataNum value="<%=Num%>">
  <INPUT type=hidden name=SortFlag value="<%=SortFlag%>" >
  <INPUT type=hidden name=SortKye value="<%=SortKye%>" >
  <INPUT type=hidden name=InfoFlag value="0">
  <INPUT type=hidden name=CONnum value="" >
  <INPUT type=hidden name=WkconNo value="" >
</Form>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
