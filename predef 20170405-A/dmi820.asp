<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi820.asp				_/
'_/	Function	:���O����oCSV���͎捞�E�o�^		_/
'_/	Date		:2003/05/30				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
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
  WriteLogH "b302", "����o���O������","06",""

'���[�U�f�[�^����
  dim USER,COMPcd,tFlag
  USER   = UCase(Session.Contents("userid"))
  COMPcd = UCase(Session.Contents("COMPcd"))

'�t�@�C�����f�[�^���擾����
  dim aryBinary,strType,nPos,nAsc,sPos,count
  dim strFile  '�t�@�C���g���q
  dim dataA    '�f�[�^

  ' �o�C�i���f�[�^���擾
  aryBinary = Request.BinaryRead(Request.TotalBytes)
  nPos=1
  count=0
  Do
    '��s���Ǎ���
    strType = ""
    'On Error Resume Next
    Do
      ' �R���e���c���擾
      nAsc = MidB(aryBinary,nPos,1)
      nAsc = AscB(nAsc)
      If (&h81 <= nAsc And nAsc <= &h9F) Or (&hE0 <= nAsc And nAsc <= &hEF) Then
        strType = strType & Chr(nAsc*256+AscB(MidB(aryBinary,nPos+1,1)))
        nPos = nPos + 1
      Else
        strType = strType & Chr(nAsc)
      End If
      If Right(strType,4) = vbCrLf & vbCrLf Then
        Exit Do
      End If
      nPos = nPos + 1
    Loop While nPos < UBound(aryBinary)
    If nPos = UBound(aryBinary) Then
      Exit Do
    End If
    If count=0 Then
      strFile = Mid(strType,InStr(LCase(strType),"filename=")+9)
      strFile = Mid(strFile,2,InStr(Mid(LCase(strFile),2),"""")-1)
      strFile = Mid(strFile,InStrRev(strFile,".")+1)
    ElseIf count=1 Then
      dataA = Split(strType, vbCrLf , -1, 1)
    End If
    count=count+1
'Response.Write strType & "<P>"
  Loop While nPos < UBound(aryBinary)

  dim ret,tmpA,ret2
  ret = true
  ret2=0
'  If strFile <> "csv" Then
'    ret=false
'    ret2=0
'    tmpA = Array("-","-")	'CW-026 ADD
'  ElseIf InStr(1,dataA(0),",",1) = 0 Then	'CW-027 ADD
  If InStr(1,dataA(0),",",1) = 0 Then	'CW-027 ADD
    ret=false					'CW-027 ADD
    ret2=1					'CW-027 ADD
    tmpA = Array("-","-")			'CW-027 ADD
  Else
    If Left(dataA(0),1)=Chr(10) OR Left(dataA(0),1)=Chr(13)Then
      dataA(0) = Mid(dataA(0),2)
    End If
    '�G���[�g���b�v�J�n
    on error resume next
    'DB�ڑ�
    dim ObjConn, ObjRS, StrSQL
    ConnDBH ObjConn, ObjRS

    dim i,CMPcd,FullName,PFlag
    CMPcd = Array("","","","","")

    For i = 0 to UBound(dataA)

    '�f�[�^�`�F�b�N
'CW-053      If tmpA(0)= "" Then	'�t�@�C���̏I��
      If Trim(dataA(i))= "" Then	'�t�@�C���̏I��
        objConn.CommitTrans
        Exit For
      End If
      tmpA = Split(UCase(Trim(dataA(i))), ",", 3, 1)
      checkStr tmpA(0), ret		'�u�b�L���O�ԍ��̃`�F�b�N
      If Not ret Then
	ret2=2
        errerF ObjConn, ObjRS, ret
        Exit For
      End If
'      If tmpA(1)="" Then
'        ret2=3
'        errerF ObjConn, ObjRS, ret
'        Exit For
'      End If
      If tmpA(1)<>"" Then
       '�w�b�h��ЃR�[�h�̃`�F�b�N
        CMPcd(1) = tmpA(1)
        checkHdCd ObjConn, ObjRS, CMPcd, ret
        If Not ret Then
          ret2=4
          errerF ObjConn, ObjRS, ret
          Exit For
        End If
        if err <> 0 then
          ObjRS.Close
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","����o�FCSV�f�[�^�o�^","102","�w�b�hID�`�F�b�N�Ɏ��s<BR>"&StrSQL
        end if
      End If
    '�u�b�N�̏d���o�^�`�F�b�N
      checkSPBook ObjConn, ObjRS, tmpA(0), PFlag, ret
      If Not ret Then
        ret2=5
        errerF ObjConn, ObjRS, ret
        Exit For
      End If
      If tmpA(1)<>"" Then		'20031112 add
    '�������^�ƎҖ��擾
        StrSQL = "SELECT FullName FROM mUsers WHERE mUsers.HeadCompanyCode='" & tmpA(1) &"'"
        ObjRS.Open StrSQL, ObjConn
        if err <> 0 then
          ObjRS.Close
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","����o�FCSV�f�[�^�o�^","102","�������^�ƎҖ��擾�Ɏ��s<BR>"&StrSQL
        end if
        FullName = ObjRS("FullName")
        ObjRS.close
      End If 				'20031112 add
    '�o�^
'CW-052 ADD Start
      If tmpA(1) = COMPcd Then 
        tFlag=1
      Else
        tFlag=0
      End If
'CW-052 ADD END

      If PFlag="0" Then
        StrSQL = "Insert Into SPBookInfo (BookNo, SenderCode, UpdtTime, UpdtPgCd, UpdtTmnl, Status,"&_
                 " Process, InputDate, TruckerCode, TruckerFlag, TruckerName ) "&_
                 "values ('"& tmpA(0) &"','"& USER &"','"& Now() &"','PREDEF01','"& USER &"','0',"&_
                 "'R','"& Now() &"','"& tmpA(1) &"','"& tFlag &"','"& FullName &"')"
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","����o�FCSV�f�[�^�o�^","103","�f�[�^�o�^�Ɏ��s<BR>"&StrSQL
        end if
      Else
        StrSQL = "UPDATE SPBookInfo SET SenderCode='"& USER &"', UpdtTime='"& Now() &"', UpdtPgCd='PREDEF01', "&_
                 "UpdtTmnl='"& USER &"', Status='0', Process='R', InputDate='"& Now() &"', "&_
                 "TruckerCode='"& tmpA(1) &"', TruckerFlag='"& tFlag &"', TruckerName='"& FullName &"' "&_
                 "WHERE BookNo='"& tmpA(0) &"' "
        ObjConn.Execute(StrSQL)
        if err <> 0 then
          Set ObjRS = Nothing
          jampErrerPDB ObjConn,"1","b302","06","����o�FCSV�f�[�^�o�^","104","�f�[�^�o�^�Ɏ��s<BR>"&StrSQL
        end if
      End If
    Next
  'DB�ڑ�����
    DisConnDBH ObjConn, ObjRS
  '�G���[�g���b�v����
    on error goto 0
  End If

  dim tmpstr
  If ret Then
    tmpstr=",���͓��e�̐���:0(������)"
  Else
    tmpstr=",���͓��e�̐���:1(���)"
  End If
  WriteLogH "b302", "����o���O������","06",tmpA(0)&"/"&tmpA(1)&tmpstr

Function checkStr(str, ret)
  dim checkChr,i,checkF
  checkChr="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ- /"
'CW-054  If Len(str) = 0 Then
  If Len(str) = 0 Or Len(str) > 21 Then
      ret = false
      Exit Function
  End If 
  For i= 1 To Len(str)
    If InStr(1,checkChr,Mid(str,i,1),1) = 0 Then
      ret = false
      Exit Function
    End If
  Next
End Function

Function errerF(ObjRS, StrSQL, ret)
  ObjConn.RollbackTrans	'���[���o�b�N
  ret = false
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���O����oCSV����</TITLE>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(600,400);
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------���O����oCSV����--------------------------->
<P><B>���O����oCSV���͏���</B></P>
<CENTER>
<%If ret Then %>
  <P><%=i%>���o�^���܂���</P>
  <INPUT type=button onClick="window.close()" value="����">
<% Else %>
<P><DIV class=alert>�G���[<P>
  <% Select Case ret2
       Case "0" %>
      �w�肳�ꂽ�t�@�C���̊g���q��CSV�ł͂���܂���B
  <%   Case "1" %>
      �w�肳�ꂽ�t�@�C����Null�܂��̓t�H�[�}�b�g���s���ł��B
  <%   Case "2" %>
      <%=i+1%>�Ԗڂ̃u�b�L���O�ԍ����s���ł��B<BR>
      Null�܂���21���ȏォ�s���ȕ����܂܂�Ă��܂��B<BR>
      �C���������x��蒼���Ă��������B
  <%   Case "3" %>
      <%=i+1%>�Ԗڂ̉�ЃR�[�h���w�肳��Ă��܂���B<BR>�C���������x��蒼���Ă��������B
  <%   Case "4" %>
      <%=i+1%>�Ԗڂ̉�ЃR�[�h�͑��݂��܂���B<BR>�C���������x��蒼���Ă��������B
  <%   Case "5" %>
      <%=i+1%>�Ԗڂ̃u�b�L���O�ԍ��͊��ɓo�^����Ă��܂��B<BR>�C���������x��蒼���Ă��������B
  <%   End Select %>
</DIV></P>
  <INPUT type=button onClick="window.history.back();" value="�߂�">
<% End If%>
</CENTER>
</BODY>
</HTML>
