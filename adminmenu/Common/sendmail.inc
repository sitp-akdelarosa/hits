<%
dim ret
ret=""
function gfSendMail(mailto, mailfrom, subject, body, mailserver)
    Dim objMail

	gfSendMail = ""
	Set objMail = CreateObject("basp21")
	if trim(mailto)<>"" and trim(mailfrom)<>"" and mailserver<>"" then		
		gfSendMail=objMail.Sendmail(mailserver, mailto, mailfrom, subject, body, "")
		if gfSendMail<>"" then
			if left(gfSendMail,3)="501" then			
				Set objMail = Nothing				
				exit function
			end if								
			gfLogMailError(gfSendMail)
		end if		
	end if
	Set objMail = Nothing
end function
function gfLogMailError(mailmsg)
   dim fo,tf,fn
   fn=Server.MapPath("/adminmenu/log/maillog.txt")
   set fo = CreateObject("Scripting.FileSystemObject")
   if fo.fileexists(fn) then
        set tf = fo.opentextfile(fn,8,True)
        if mailmsg<>"" then
            tf.Write mailmsg & vbNewLine
            tf.close
            set tf= Nothing
            exit function
        end if
   end if
end function
%>