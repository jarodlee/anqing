<script language="JavaScript">
/*
Copyright ?MaXimuS 1999-2000, All Rights Reserved.
Site: http://www.absolutegb.com/maximus
E-mail: maximus@nsimail.com
*/
<!--
var MyDocObject
IE = (document.all)
NS = (navigator.appName=="Netscape" && navigator.appVersion >= "4")
NS6 = (parseInt(navigator.appVersion) >= 5)
isMac = (navigator.appVersion.indexOf("Mac")!=-1) ? true : false; 

/*
Configure menu styles below
NOTE: To edit the link colors, go to the STYLE tags and edit the ssm2Items colors
*/
hdrFontFamily="<%= strDefaultFontFace %>";
hdrFontSize="<%= strDefaultFontSize %>";
hdrFontColor='<%= strCategoryFontColor %>';
hdrBGColor="<%= strcategoryCellColor %>";
linkFontFamily='<%= strDefaultFontFace %>';
linkFontSize="<%= strDefaultFontSize %>";
linkBGColor='<%= strForumCellColor %>';
linkOverBGColor='<%= strAltForumCellColor %>';
linkTarget="_top";
YOffset=<%= strYOffset %>;
staticYOffset=20;
menuBGColor='<%= strTableBorderColor %>';
menuIsStatic="yes";
menuHeader="论坛会员"
menuWidth=<%= strMenuWidth %>; // Must be a multiple of 10!
staticMode="advanced"
barBGColor='<%= strHeadCellColor %>';
barFontFamily='<%= strDefaultFontFace %>';
barFontSize='<%= strDefaultFontSize %>';
barFontColor='<%= strHeadFontColor %>';
barText="论坛菜单";

function moveOut() {
if (window.cancel) {cancel="";}
if (window.moving2) {clearTimeout(moving2); moving2="";}
if ((IE && ssm2.style.pixelLeft<0)||(NS && document.ssm2.left<0)) 
{
	if (IE) {ssm2.style.pixelLeft += 10;}
	if (NS) {document.ssm2.left += 10;}
	moving1 = setTimeout('moveOut()', 10)
}
else {clearTimeout(moving1)}
};

function moveBack() {
cancel = moveBack1()
}
function moveBack1() {
if (window.moving1) {clearTimeout(moving1)}
if ((IE && ssm2.style.pixelLeft>(-menuWidth))||(NS && document.ssm2.left>(-menuWidth))) {
	if (IE) {ssm2.style.pixelLeft -= 10;}
	if (NS) {document.ssm2.left -= 10;}
	moving2 = setTimeout('moveBack1()', 10)
}
else {clearTimeout(moving2)}
};

lastY = 0;
function makeStatic(mode) {
if (IE) {winY = document.body.scrollTop;var NM=ssm2.style}
if (NS) {winY = window.pageYOffset;var NM=document.ssm2}
if (NS6) {winY = window.pageYOffset;var NM=document.getElementById('ssm2').style}
if (mode=="smooth") {
	if ((IE||NS) && winY!=lastY) 
	{
		smooth = .2 * (winY - lastY);
		if(smooth > 0) 
			smooth = Math.ceil(smooth);
		else smooth = Math.floor(smooth);
		if (IE) NM.pixelTop+=smooth;
		if (NS) NM.top+=smooth;
		lastY = lastY+smooth;
	}
	setTimeout('makeStatic("smooth")', 1)
}
else if (mode=="advanced") 
{
	if ((IE||NS) && winY>YOffset-staticYOffset) {
		if (IE) {NM.pixelTop=winY+staticYOffset}
		if (NS) {NM.top=winY+staticYOffset}
	}
	else {
		if (IE) {NM.pixelTop=YOffset}
		if (NS) {NM.top=YOffset-7}
	}
	setTimeout('makeStatic("advanced")', 1)}
}

function init() {
if (isMac) return;
if (NS6)  {
document.getElementById('ssm2').style.visibility = "visible";
document.getElementById('ssm2').style.Left = -menuWidth;
document.getElementById('ssm2').style.Left = YOffset;
}
if (IE) {
ssm2.style.pixelLeft = -menuWidth;
ssm2.style.visibility = "visible"}
else if (NS) {
document.ssm2.left = -menuWidth;
document.ssm2.visibility = "show"}
else {alert('Choose either the "smooth" or "advanced" static modes!')}}

function MoveTo(x,y) 
{  
	var obj = document.getElementById('ssm2')
   
	obj.style.left = x; 

}
-->
</script>