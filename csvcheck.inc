<% 
'/*
' *	�e���ڌ������������`�F�b�N
' *
' *	@param	String	strAnyTmp	���ڕ�����
' *	@param	String	strName		���ږ�
' *	@param	int		iFig		����
' *	@param	int		iDateFlag	���t���ڂ��ǂ����i���t�ł͂Ȃ��F0 YYYYMMDDHHMM�F1 YYYYMMDD�F2�j
' *	@param	boolean	bRecFlag	�K�{���ڂ��ǂ����i�K�{�Ftrue�j
' *	@param	boolean	bNumFlag	���l���ڂ��ǂ����i���l�Ftrue�j
' *	@return	String				�G���[���b�Z�[�W
' */
Function CheckParam(strAnyTmp,strName,iFig,iDateFlag,bRecFlag,bNumFlag) 

	Dim strError
	strAnyTmp = Trim(strAnyTmp)
	strError=""

	'����
	If Len(strAnyTmp)>iFig Then
		strError=strError & strName & "�̒������ُ�ł��B(" & strAnyTmp & ")<BR> "
	End If
	'�K�{
	If bRecFlag And Len(strAnyTmp)<=0 Then
		strError=strError & strName & "�͕K�{���ڂł��B(" & strAnyTmp & ")<BR> "
	End If
	'����
	If Len(strAnyTmp)<>0 And bNumFlag And Not IsNumeric(strAnyTmp) Then
		strError=strError & strName & "�ɂ͐��l���L�����ĉ������B(" & strAnyTmp & ")<BR> "
	End If
	'���t
	If Len(strAnyTmp)<>0 And iDateFlag=1 And Len(strAnyTmp)<>12 Then
		strError=strError & strName & "�ɂ� ""/"" �ŋ�؂������l���L�����ĉ������B(" & strAnyTmp & ") <BR>�i��j2002�N1��23�� �ߌ�4��5�� => 2002/1/23/16/05<BR> "
	End If
	If Len(strAnyTmp)<>0 And iDateFlag=2 And Len(strAnyTmp)<>8 Then
		strError=strError & strName & "�ɂ� ""/"" �ŋ�؂������l���L�����ĉ������B(" & strAnyTmp & ") <BR>�i��j2002�N1��23�� => 2002/1/23<BR> "
	End If

	CheckParam = strError

End Function

%>