<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo180.asp				_/
'_/	Function	:��������ꗗCSV�t�@�C���_�E�����[�h	_/
'_/	Date		:2003/07/31				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<!--#include File="./Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b208", "��������O���CSV�t�@�C���_�E�����[�h","01",""


'�f�[�^�擾
  dim Num,DtTbl,i,j
  Get_Data Num,DtTbl

' �t�@�C���̃_�E�����[�h
  Response.ContentType="application/octet-stream"
  Response.AddHeader "Content-Disposition","attachment; filename=output.csv"
    Response.Write "�����\���,�w����,�w�����ւ̉�,�R���e�i�ԍ�,�D��,�D��,�T�C�Y,�ԋp��,"
    Response.Write "�f�B�e���V�����t���[�^�C��,�w����,�w�����,���l"
    Response.Write Chr(13) & Chr(10)
    For j=1 To Num
      Response.Write Trim(DtTbl(j)(1))&","&Trim(DtTbl(j)(2))&","&Trim(DtTbl(j)(8))&","&Trim(DtTbl(j)(3))&","
      Response.Write DtTbl(j)(9)&","&Trim(DtTbl(j)(10))&","&Trim(DtTbl(j)(11))&","&Trim(DtTbl(j)(12))&","
      Response.Write Trim(DtTbl(j)(13))&","&Trim(DtTbl(j)(5))&","&Trim(DtTbl(j)(6))&","&Trim(DtTbl(j)(14))&","
      Response.Write Chr(13) & Chr(10)
    Next
  Response.End

%>
