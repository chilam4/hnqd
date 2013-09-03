<%
dim AdminAction
AdminAction=request.QueryString("AdminAction")
select case AdminAction
  case "Out"
    call OutLogin()
  case else
    call Login()
end select
sub Login()
  if session("AdminName")="" Or session("UserName")="" Or session("AdminPurview")="" Or session("LoginSystem")<>"Succeed"  Or (EnableSiteManageCode = True And Trim(Request.Cookies("AdminLoginCode")) <> SiteManageCode) then
	 response.redirect "Admin_Login.asp"
     response.end
  end if
end sub
sub OutLogin()
  session.contents.remove "AdminName"
  session.contents.remove "UserName"
  session.contents.remove "AdminPurview"
  session.contents.remove "LoginSystem"
  session.contents.remove "VerifyCode"
  response.redirect "Admin_Login.asp"
end sub
%>