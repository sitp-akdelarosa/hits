<% 
rem ���ϐ�"LOGON_USER"�̓��e���烆�[�UID�����o��
Function GetUserID(sID) 
	ptr = len(sID)
	do while 0 < ptr 
		if mid(sID, ptr, 1) = "\" then
			exit do
		end if
		ptr = ptr - 1
	loop
	if ptr = 0 then
		GetUserID = sID
	else
		if ptr = len(sID) then
			GetUserID = ""
		else
			GetUserID = left(right(sID, len(sID) - ptr),8)
		end if
	end if
End Function
%>