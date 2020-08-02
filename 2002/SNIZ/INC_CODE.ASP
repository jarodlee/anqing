<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  汉化修改: 资源搜罗站                                         #
'#  电子邮件: cgier@21cn.com                                     #
'#  主页地址: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  论坛地址:http://ubb.yesky.net                                #
'#  最后修改日期: 2001/03/12    中文版本：Version 3.1 SR4        #
'#################################################################
'# 原始来源                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#【版权声明】                                                   #
'#                                                               #
'# 本软体为共享软体(shareware)提供个人网站免费使用，请勿非法修改,#
'# 转载，散播，或用于其他图利行为，并请勿删除版权声明。          #
'# 如果您的网站正式起用了这个脚本，请您通知我们，以便我们能够知晓#
'# 如果可能，请在您的网站做上我们的链接，希望能给予合作。谢谢！  #
'#################################################################
'# 请您尊重我们的劳动和版权，不要删除以上的版权声明部分，谢谢合作#
'# 如有任何问题请到我们的论坛告诉我们                            #
'#################################################################
%>
<SCRIPT LANGUAGE="JavaScript">

helpstat = false;
stprompt = false;
basic = true;


function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 0) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 2) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}

function AddText(NewCode) {
document.PostTopic.Message.value+=NewCode
}

function email() {
	if (helpstat) {
		alert("Email标签\n将一个Email地址转换成Email链接。\n用法一 : [url]someone\@anywhere.com[/url] \n用法二 : [url=\"someone\@anywhere.com\"]Email链接信息[/url]");
		}
	else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
		}
	else { 
		txt2=prompt("输入在Email链接中想要显示的信息，如：情哥的邮件地址\n您若希望直接显示Email地址，就空着不要填写，然后点确定按钮！",""); 
		if (txt2!=null) {
			txt=prompt("请输入Email地址.","mailto:");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt+"[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url=\""+txt+"\"]"+txt2+"[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}
function showsize(size) {
	if (helpstat) {
		alert("字体大小标签\n设置字体大小。\n可选值为1号到6号。\n 1号是最小字体，6号是最大字体。.\n\n法: [size="+size+"]设置为 "+size+" 号字体的文本[/size="+size+"]");
	} else if (basic) {
		AddTxt="[size="+size+"][/size="+size+"]";
		AddText(AddTxt);
	} else {                       
		txt=prompt("设置字体大小为 "+size,"输入需要设置字体大小的文本"); 
		if (txt!=null) {             
			AddTxt="[size="+size+"]"+txt+"[/size="+size+"]";
			AddText(AddTxt);
		}        
	}
}

function bold() {
	if (helpstat) {
		alert("粗体标签.\n将文本字体加粗显示..\n\n用法: [b]需要加粗显示的文本[/b]");
	} else if (basic) {
		AddTxt="[b][/b]";
		AddText(AddTxt);
	} else {  
		txt=prompt("加粗显示.","输入想要加粗显示的文本");     
		if (txt!=null) {           
			AddTxt="[b]"+txt+"[/b]";
			AddText(AddTxt);
		}       
	}
}

function italicize() {
	if (helpstat) {
		alert("斜体标签\n将文本字体设置为斜体.\n\n用法: [i]需要设置为斜体的文本[/i]");
	} else if (basic) {
		AddTxt="[i][/i]";
		AddText(AddTxt);
	} else {   
		txt=prompt("设置为斜体","输入需要设置为斜体的文本");     
		if (txt!=null) {           
			AddTxt="[i]"+txt+"[/i]";
			AddText(AddTxt);
		}	        
	}
}

function quote() {
	if (helpstat){
		alert("引用标签\n引用某人发表的一段文字.\n\n用法: [quote]需引用的内容[/quote]");
	} else if (basic) {
		AddTxt=" [quote] [/quote]";
		AddText(AddTxt);
	} else {   
		txt=prompt("引用","输入需要引用的内容");     
		if(txt!=null) {          
			AddTxt=" [quote] "+txt+" [/quote]";
			AddText(AddTxt);
		}	        
	}
}

function showcolor(color) {
	if (helpstat) {
		alert("文字颜色标签\n设置文字颜色.\n\n用法: ["+color+"]需要设置成 "+color+" 的文本[/"+color+"]");
	} else if (basic) {
		AddTxt="["+color+"][/"+color+"]";
		AddText(AddTxt);
	} else {  
     	txt=prompt("文字颜色设置成 "+color,"输入指定颜色的文本");
		if(txt!=null) {
			AddTxt="["+color+"]"+txt+"[/"+color+"]";
			AddText(AddTxt);        
		} 
	}
}

function center() {
 	if (helpstat) {
		alert("居中标签\n将某段文本居中显示.\n\n用法: [center]需居中显示的文本[/center]");
	} else if (basic) {
		AddTxt="[center][/center]";
		AddText(AddTxt);
	} else {  
		txt=prompt("文本居中显示","输入需居中显示的文本");     
		if (txt!=null) {          
			AddTxt="[center]"+txt+"[/center]";
			AddText(AddTxt);
		}	       
	}
}

function hyperlink() {
	if (helpstat) {
		alert("超级链接标签\n将网址转换为超级链接。.\n\n用法一: [url]http://www.sdsea.com[/url]\n\n用法二: [url=http://www.sdsea.com]超级链接信息[/url]");
	} else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("需显示的超级链接信息.\n您若希望直接显示网址，就空着不要填写，然后按确定按钮.",""); 
		if (txt2!=null) {
			txt=prompt("请输入超级链接的网址.","http://");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt+"[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url=\""+txt+"\"]"+txt2+"[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function image() {
	if (helpstat){
		alert("图片标签\n在帖子中插入一个图片。.\n\n用法: [img]http://www.sdsea.com/logo.gif[/img]");
	} else if (basic) {
		AddTxt="[img][/img]";
		AddText(AddTxt);
	} else {  
		txt=prompt("请输入图片链接的网址","http://");    
		if(txt!=null) {            
			AddTxt="[img]"+txt+"[/img]";
			AddText(AddTxt);
		}	
	}
}

function showcode() {
	if (helpstat) {
		alert("代码标签\n在输入的信息头尾加上水平线，使得输入信息保持原来的格式。\n主要用于以正确的格式显示一些代码，如Perl、javascripts代码。.\n\n用法: [code]需要以原始格式显示的代码内容[/code]");
	} else if (basic) {
		AddTxt=" [code] [/code]";
		AddText(AddTxt);
	} else {   
		txt=prompt("输入需要以原始格式显示的代码内容","");     
		if (txt!=null) {          
			AddTxt="[code]"+txt+"[/code]";
			AddText(AddTxt);
		}	       
	}
}

function list() {
	if (helpstat) {
		alert("编号标签\n创建一个以园点、数字或者字母为编号的列表.\n\n用法: [list] [*]项目一[/*] [*]项目二[/*] [*]项目三[/*] [/list]");
	} else if (basic) {
		AddTxt=" [list][*]  [/*][*]  [/*][*]  [/*][/list]";
		AddText(AddTxt);
	} else {  
		type=prompt("编号类型\n输入\'A\' 显示为字母编号, \'1\' 显示为数字编号，空白显示为园点编号.","");               
		while ((type!="") && (type!="A") && (type!="a") && (type!="1") && (type!=null)) {
			txt=prompt("错误！\\n编号类型只能是空白，或者输入 'A' 和 '1'.","");               
		}
		if (type!=null) {
			if (type=="") {
				AddTxt="[list]";
			} else {
				AddTxt="[list="+type+"]";
			} 
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("空白表示结束列表",""); 
				if (txt!="") {             
					AddTxt+="[*]"+txt+"[/*]"; 
				}                   
			} 
			if (type=="") {
				AddTxt+="[/list] ";
			} else {
				AddTxt+="[/list="+type+"]";
			} 
			AddText(AddTxt); 
		}
	}
}

function underline() {
  	if (helpstat) {
		alert("下划线标签\n将指定文字加上下划线\n\n用法: [u]需加上下划线的文本[/u]");
	} else if (basic) {
		AddTxt="[u][/u]";
		AddText(AddTxt);
	} else {  
		txt=prompt("输入需加上下划线的文本.","输入需加上下划线的文本");     
		if (txt!=null) {           
			AddTxt="[u]"+txt+"[/u]";
			AddText(AddTxt);
		}	        
	}
}

function showfont(font) {
 	if (helpstat){
		alert("字体标签\n设置文本字体\n字体能否正确显示取决于浏览者电脑中是否安装指定字体\n因此，中文字体推荐宋体和黑体这两种常用字体.\n\n用法: [font="+font+"]指定文本的字体为 "+font+"[/font]");
	} else if (basic) {
		AddTxt="[font="+font+"][/font="+font+"]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("文本的字体为 "+font,"输入需设置字体的文本");
		if (txt!=null) {             
			AddTxt="[font="+font+"]"+txt+"[/font="+font+"]";
			AddText(AddTxt);
		}        
	}  
}
</SCRIPT>