<script language="javascript">
function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
</script>
<script Language="JavaScript">
//***********Ĭ�����ö���.*********************
tPopWait=50;//ͣ��tWait�������ʾ��ʾ��
tPopShow=5000;//��ʾtShow�����ر���ʾ
showPopStep=20;
popOpacity=99;

//***************�ڲ���������*****************
sPop=null;
curShow=null;
tFadeOut=null;
tFadeIn=null;
tFadeWaiting=null;

document.write("<style type='text/css'id='defaultPopStyle'>");
document.write(".cPopText {  background-color: #F8F8F5;color:#000000; border: 1px #000000 solid;font-color: font-size: 12px; padding-right: 4px; padding-left: 4px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0)}");
document.write("</style>");
document.write("<div id='dypopLayer' style='position:absolute;z-index:1000;' class='cPopText'></div>");


function showPopupText(){
var o=event.srcElement;
	MouseX=event.x;
	MouseY=event.y;
	if(o.alt!=null && o.alt!=""){o.dypop=o.alt;o.alt=""};
        if(o.title!=null && o.title!=""){o.dypop=o.title;o.title=""};
	if(o.dypop!=sPop) {
			sPop=o.dypop;
			clearTimeout(curShow);
			clearTimeout(tFadeOut);
			clearTimeout(tFadeIn);
			clearTimeout(tFadeWaiting);	
			if(sPop==null || sPop=="") {
				dypopLayer.innerHTML="";
				dypopLayer.style.filter="Alpha()";
				dypopLayer.filters.Alpha.opacity=0;	
				}
			else {
				if(o.dyclass!=null) popStyle=o.dyclass 
					else popStyle="cPopText";
				curShow=setTimeout("showIt()",tPopWait);
			}
			
	}
}

function showIt(){
		dypopLayer.className=popStyle;
		dypopLayer.innerHTML=sPop;
		popWidth=dypopLayer.clientWidth;
		popHeight=dypopLayer.clientHeight;
		if(MouseX+12+popWidth>document.body.clientWidth) popLeftAdjust=-popWidth-24
			else popLeftAdjust=0;
		if(MouseY+12+popHeight>document.body.clientHeight) popTopAdjust=-popHeight-24
			else popTopAdjust=0;
		dypopLayer.style.left=MouseX+12+document.body.scrollLeft+popLeftAdjust;
		dypopLayer.style.top=MouseY+12+document.body.scrollTop+popTopAdjust;
		dypopLayer.style.filter="Alpha(Opacity=0)";
		fadeOut();
}

function fadeOut(){
	if(dypopLayer.filters.Alpha.opacity<popOpacity) {
		dypopLayer.filters.Alpha.opacity+=showPopStep;
		tFadeOut=setTimeout("fadeOut()",1);
		}
		else {
			dypopLayer.filters.Alpha.opacity=popOpacity;
			tFadeWaiting=setTimeout("fadeIn()",tPopShow);
			}
}

function fadeIn(){
	if(dypopLayer.filters.Alpha.opacity>0) {
		dypopLayer.filters.Alpha.opacity-=1;
		tFadeIn=setTimeout("fadeIn()",1);
		}
}
document.onmouseover=showPopupText;

function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }

//�����˵���ش���
 var h;
 var w;
 var l;
 var t;
 var topMar = 1;
 var leftMar = -2;
 var space = 1;
 var isvisible;
 var MENU_SHADOW_COLOR='#999999';//���������˵���Ӱɫ
 var global = window.document
 global.fo_currentMenu = null
 global.fo_shadows = new Array

function HideMenu() 
{
 var mX;
 var mY;
 var vDiv;
 var mDiv;
	if (isvisible == true)
{
		vDiv = document.all("menuDiv");
		mX = window.event.clientX + document.body.scrollLeft;
		mY = window.event.clientY + document.body.scrollTop;
		if ((mX < parseInt(vDiv.style.left)) || (mX > parseInt(vDiv.style.left)+vDiv.offsetWidth) || (mY < parseInt(vDiv.style.top)-h) || (mY > parseInt(vDiv.style.top)+vDiv.offsetHeight)){
			vDiv.style.visibility = "hidden";
			isvisible = false;
		}
}
}

function ShowMenu(vMnuCode,tWidth) {
	vSrc = window.event.srcElement;
	vMnuCode = "<table id='submenu' cellspacing=1 cellpadding=3 style='width:"+tWidth+"' class=tableborder1 onmouseout='HideMenu()'><tr height=23><td nowrap align=left class=tablebody1>" + vMnuCode + "</td></tr></table>";

	h = vSrc.offsetHeight;
	w = vSrc.offsetWidth;
	l = vSrc.offsetLeft + leftMar+4;
	t = vSrc.offsetTop + topMar + h + space-2;
	vParent = vSrc.offsetParent;
	while (vParent.tagName.toUpperCase() != "BODY")
	{
		l += vParent.offsetLeft;
		t += vParent.offsetTop;
		vParent = vParent.offsetParent;
	}

	menuDiv.innerHTML = vMnuCode;
	menuDiv.style.top = t;
	menuDiv.style.left = l;
	menuDiv.style.visibility = "visible";
	isvisible = true;
    makeRectangularDropShadow(submenu, MENU_SHADOW_COLOR, 4)
}

function makeRectangularDropShadow(el, color, size)
{
	var i;
	for (i=size; i>0; i--)
	{
		var rect = document.createElement('div');
		var rs = rect.style
		rs.position = 'absolute';
		rs.left = (el.style.posLeft + i) + 'px';
		rs.top = (el.style.posTop + i) + 'px';
		rs.width = el.offsetWidth + 'px';
		rs.height = el.offsetHeight + 'px';
		rs.zIndex = el.style.zIndex - i;
		rs.backgroundColor = color;
		var opacity = 1 - i / (i + 1);
		rs.filter = 'alpha(opacity=' + (100 * opacity) + ')';
		el.insertAdjacentElement('afterEnd', rect);
		global.fo_shadows[global.fo_shadows.length] = rect;
	}
}
//�û��������
var manage= '<a style=font-size:9pt;line-height:14pt; href=\"JavaScript:openScript(\'messanger.asp?action=new\',500,400)\">������</a><br><a style=font-size:9pt;line-height:14pt; href=\"dispuser.asp?id=<%=userid%>&boardid=<%=boardid%>&action=permission\">������ʲô</a><br><a style=font-size:9pt;line-height:14pt; href=\"topicwithme.asp?s=2\">�ҷ��������</a><br><a style=font-size:9pt;line-height:14pt; href=\"topicwithme.asp?s=1\">�Ҳ��������</a><br><a style=font-size:9pt;line-height:14pt; href=\"mymodify.asp\">���������޸�</a><br><a style=font-size:9pt;line-height:14pt; href=\"modifypsw.asp\">�û������޸�</a><br><a style=font-size:9pt;line-height:14pt; href=\"modifyadd.asp\">��ϵ�����޸�</a><br><a style=font-size:9pt;line-height:14pt; href=\"usersms.asp\">�û����ŷ���</a><br><a style=font-size:9pt;line-height:14pt; href=\"friendlist.asp\">�༭�����б�</a><br><a style=font-size:9pt;line-height:14pt; href=\"favlist.asp\">�û��ղع���</a><br><a style=font-size:9pt;line-height:14pt; href=\"myfile.asp\">�����ļ�����</a>'
//ģ���б�
var stylelist = '<%
myCache.name="stylelist"
if myCache.valid then
response.write replace(myCache.value,"{boardid}",boardid)

else
dim stylelist
stylelist="<a style=font-size:9pt;line-height:12pt; href=\""cookies.asp?action=stylemod&skinid=0&boardid={boardid}\"">�ָ�Ĭ������</a><br>"
set rs=conn.execute("select id,skinname from config order by id")
do while not rs.eof
stylelist=stylelist & "<a style=font-size:9pt;line-height:12pt; href=\""cookies.asp?action=stylemod&skinid="&rs(0)&"&boardid={boardid}\"">"&rs(1)&"</a><br>"
rs.movenext
loop
myCache.add stylelist,dateadd("n",9999,now)
response.write replace(stylelist,"{boardid}",boardid)
end if
%>'
//��̳״̬
var boardstat= '<a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?boardid=<%=boardid%>\">��������ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?action=lasttopicnum&boardid=<%=boardid%>\">������ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?action=lastbbsnum&boardid=<%=boardid%>\">������ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?reaction=online&boardid=<%=boardid%>\">����ͼ��</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?reaction=onlineinfo&boardid=<%=boardid%>\">�������</a><br><a style=font-size:9pt;line-height:14pt; href=\"boardstat.asp?reaction=onlineUserinfo&boardid=<%=boardid%>\">�û�������ͼ��</a>'
//��̳�ղ�
var downlist= '<a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=0&boardid=<%=boardid%>\">�ļ������</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=1&boardid=<%=boardid%>\">ͼƬ�����</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=2&boardid=<%=boardid%>\">Flash���</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=3&boardid=<%=boardid%>\">���ּ����</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp?filetype=4&boardid=<%=boardid%>\">��Ӱ�����</a><br><a style=font-size:9pt;line-height:14pt; href=\"show.asp\">�ؿ�����</a>'
</script>