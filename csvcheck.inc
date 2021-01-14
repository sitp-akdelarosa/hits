<% 
'/*
' *	各項目桁数＆整合性チェック
' *
' *	@param	String	strAnyTmp	項目文字列
' *	@param	String	strName		項目名
' *	@param	int		iFig		桁数
' *	@param	int		iDateFlag	日付項目かどうか（日付ではない：0 YYYYMMDDHHMM：1 YYYYMMDD：2）
' *	@param	boolean	bRecFlag	必須項目かどうか（必須：true）
' *	@param	boolean	bNumFlag	数値項目かどうか（数値：true）
' *	@return	String				エラーメッセージ
' */
Function CheckParam(strAnyTmp,strName,iFig,iDateFlag,bRecFlag,bNumFlag) 

	Dim strError
	strAnyTmp = Trim(strAnyTmp)
	strError=""

	'桁数
	If Len(strAnyTmp)>iFig Then
		strError=strError & strName & "の長さが異常です。(" & strAnyTmp & ")<BR> "
	End If
	'必須
	If bRecFlag And Len(strAnyTmp)<=0 Then
		strError=strError & strName & "は必須項目です。(" & strAnyTmp & ")<BR> "
	End If
	'数字
	If Len(strAnyTmp)<>0 And bNumFlag And Not IsNumeric(strAnyTmp) Then
		strError=strError & strName & "には数値を記入して下さい。(" & strAnyTmp & ")<BR> "
	End If
	'日付
	If Len(strAnyTmp)<>0 And iDateFlag=1 And Len(strAnyTmp)<>12 Then
		strError=strError & strName & "には ""/"" で区切った数値を記入して下さい。(" & strAnyTmp & ") <BR>（例）2002年1月23日 午後4時5分 => 2002/1/23/16/05<BR> "
	End If
	If Len(strAnyTmp)<>0 And iDateFlag=2 And Len(strAnyTmp)<>8 Then
		strError=strError & strName & "には ""/"" で区切った数値を記入して下さい。(" & strAnyTmp & ") <BR>（例）2002年1月23日 => 2002/1/23<BR> "
	End If

	CheckParam = strError

End Function

%>