<script language="JavaScript">
function openPollWindow(url) {
  popupWin = window.open(url,'new_page','width=200,height=400')
}

if (IE) {document.write('<DIV ID="ssm2" onmouseover="moveOut()" onmouseout="moveBack()">')}
if (NS6) {document.write('<DIV ID="ssm2" onmouseover="MoveTo(5,10)" onmouseout="MoveTo('+(-menuWidth-20)+',10)">')}
else if (NS) {document.write('<LAYER visibility="hide" top="'+YOffset+'" name="ssm2" bgcolor="'+linkBGColor+'" left="0" onmouseover="moveOut()" onmouseout="moveBack()">')}
tempBar=""
for (i=0;i<barText.length;i++) {
tempBar+=barText.substring(i, i+1)+"<BR>"}
document.write('<table border="0" cellpadding="3" cellspacing="1" width="'+(menuWidth+16+2)+'" bgcolor="'+menuBGColor+'"><tr ><td align="center" bgcolor="'+hdrBGColor+'" WIDTH="'+menuWidth+'">&nbsp;<font face="'+hdrFontFamily+'" Size="'+hdrFontSize+'" COLOR="'+hdrFontColor+'"><b>'+menuHeader+'</b></font></td><td align="center" rowspan="100" width="16" bgcolor="'+barBGColor+'"><p align="center"><font face="'+barFontFamily+'" Size="'+barFontSize+'" COLOR="'+barFontColor+'"><B>'+tempBar+'</B></font></p></TD></tr>')
function addItem(text, link, target) {
if (!target) {target=linkTarget}
document.write('<TR bgcolor="'+menuBGColor+'" ><TD align="center" BGCOLOR="'+linkBGColor+'" onmouseover="bgColor=\''+linkOverBGColor+'\'" onmouseout="bgColor=\''+linkBGColor+'\'"><ILAYER><LAYER onmouseover="bgColor=\''+linkOverBGColor+'\'" onmouseout="bgColor=\''+linkBGColor+'\'" WIDTH="100%" height="100%"><FONT color="' + hdrFontColor + '" face="'+linkFontFamily+'" Size="'+linkFontSize+'">&nbsp;<A HREF="'+link+'" target="'+target+'" CLASS="ssm2Items">'+text+'</LAYER></ILAYER></TD></TR>')}
function addHdr(text) {
document.write('<tr bgcolor="'+menuBGColor+'"><td align="center" bgcolor="'+hdrBGColor+'">&nbsp;<font face="'+hdrFontFamily+'" Size="'+hdrFontSize+'" COLOR="'+hdrFontColor+'"><b>'+text+'</b></font></td></tr>')}

//Only edit the script between HERE

addItem('论坛首页','default.asp');
if ('<%= strUseExtendedProfile %>')
    		addItem('个人资料','pop_profile.asp?mode=Edit')
else
    		addItem('个人资料',"javascript:openWindow3('pop_profile.asp?mode=Edit')")
if ('<%= strAutoLogon %>')
    		addItem('会员注册','policy.asp')
addItem('最新文章','active.asp');
addItem('会员列表','members.asp');
addItem('个人日记','events.asp');

addItem('收藏夹','bookmark.asp');
if ('<%= request.serverVariables("PATH_INFO") %>' == '<%= strCookieURL %>' + 'topic.asp')
addItem('加入主题到收藏夹',"bookmark.asp?mode=add&id='<%=Request.querystring("Topic_ID")%>'")
addItem("在线会员",'active_users.asp');
if ('<%= mlev %>' == 0)
{
	if ('<%=strEmail  %>' == "1")
		addItem('忘记你的密码？',"javascript:openWindow3('pop_pword.asp')");
	if ('<%= strNoCookies %>' == "1")
		addItem('管理选项','admin_home.asp');
}
addItem('论坛搜索','search.asp');
document.write('<tr><td bgcolor="'+'<% =strForumCellColor %>'+'">')
document.write ('<table border=0 cellspacing=0 cellpadding=0>')
document.write ('	<form action="Search.asp?mode=DoIt" method="post" name="QuickSearch" target="_top">')
document.write ('			<tr><font face="'+'<% =strDefaultFontFace %>'+'" size='+'<% =strFooterFontSize %>'+'><b>快速搜索</b></font>')
document.write ('				<td>')
document.write ('					<input type="text" name="Search" size="12">')
document.write ('					<a href="JavaScript:document.QuickSearch.submit()">')
document.write ('					<img src="<%=strImageURL %>icon_quicksearch.gif" width="18" height="18" border="0" alt=""></a></td>')
document.write ('			</tr>')
document.write ('			<input type="HIDDEN" name="andor" value="and">')
document.write ('			<input type="HIDDEN" name="SearchDate" value="30">')
document.write ('		<tr><td>')
document.write ('		</td></tr></form></table>')
document.write('</td></tr>')
addItem('帮助说明','faq.asp');

//Admin Only Stuff Below

if ('<%= mlev %>' == "4")
{
addHdr('管理员专用');
addItem('管理选项','admin_home.asp');
addItem('模组管理','admin_mods.asp');
addItem('公告管理','admin_announce_home.asp');
addItem('上传文件',"javascript:openWindow3('admin_upload.asp')");
addItem('主题统计','topic_stats.asp');
}

// and HERE! No more!

document.write('</table>')
if (IE) {document.write('</DIV>')}
if (NS6) {document.write('</DIV>')}
else if (NS) {document.write('</LAYER>')}
if ((IE||NS) && (menuIsStatic=="yes"&&staticMode)) {makeStatic(staticMode);}
</script>
