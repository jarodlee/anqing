<html>
<HEAD>

<script language="PerlScript" RUNAT=SERVER>
sub viewSource(\$){
	use LWP::Simple;
	my ($URL) =  @_;
	$Source = get($URL);
	if($Source ne "")
	{    
	  	return true;  
	} 
	else 
	{
	   return false;
	}
}
</script>
</HEAD>

<BODY>

<center>
Type in a full URL and click "Validate"
<br>
<br>
<form name=getSource action="http_asp.asp?url=<%= Request.Form("url") %>">
<input type=text name=url value="<%= Request.Form("url") %>">
<br>
<br>
<input type=submit name=view value="Validate">
</form>
</center>
<p><center>
<font face="arial, helvetica" size="-2">URL Validator by Huw Reddick<br>
</center><p>

<% 	if Request.querystring("url") <> "" then
		if viewSource(Request.querystring("url")) then 
		response.write(Request.querystring("url"))
		else
		response.write("not found")
		end if
	End If %>
</font>
</body>
</html>


