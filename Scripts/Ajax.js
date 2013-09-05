/*
本段声明不会影响您程序的响应速度,请不要删除本段声明!
官方主页:http://17show.net/book/
程序开发:骑着蜗牛去飞翔(Leo/Zero Crystal)
版权所有:Copyright 2003-2005 一起秀网络(17show.net).INC All Right Reserved 
 */
var Url="";
var Sort="00000001";
var PageRecord=6;

function Trim(){
	return this.replace(/\s+$|^\s+/g,"");
}
String.prototype.Trim=Trim;	//过滤两端空格

function GetObject(elementId) { 	//获取指定id的object
	if (document.getElementById) { 
		return document.getElementById(elementId); 
	} else if (document.all) { 
		return document.all[elementId]; 
	} else if (document.layers) { 
		return document.layers[elementId]; 
	} 
}

function GetObjValue(elementId){	//获取指定id的form组件的值
	if(GetObject(elementId).value!=undefined)
		return GetObject(elementId).value.Trim();
	else
		return "";
}

function ObjXML(){
	var ObjXml;
	try{
		ObjXml=new XMLHttpRequest();
	}catch(e){
    		var a=['MSXML2.XMLHTTP.5.0','MSXML2.XMLHTTP.4.0','MSXML2.XMLHTTP.3.0','MSXML2.XMLHTTP','MICROSOFT.XMLHTTP.1.0','MICROSOFT.XMLHTTP.1','MICROSOFT.XMLHTTP'];
    		for (var i=0;i<a.length;i++){
      			try{
        			ObjXml=new ActiveXObject(a[i]);
        			break;
      			}catch(e){}
    		}
  	}
	return ObjXml;
}

function checkForm(){	//表单的检测
	if(GetObjValue("SfMess_content")==""||GetObjValue("SfMess_phone")==""||GetObjValue("SfMess_email")==""||GetObjValue("SfMess_phone")==""){
		alert("请您将主题/昵称/内容填写完整再提交！");
		return false;
	}
	if(GetObjValue("NickName").length>20){
		alert("你的名字也太长了吧？不要大于20哦");
		return false;
	}
	if(GetObjValue("QQ")!="" && !/^[0-9]{5,10}$/.test(GetObjValue("qq"))){
		alert("腾讯只有5-10位的qq号吧？");
		return false;
	}
	if(GetObjValue("Email")!="" && !/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/.test(GetObjValue("email"))){
		alert("邮箱地址错误！");
		return false;
	}
	return true;
}

function ObjXmlSend() {	//发送留言
	var ObjXml=ObjXML();
	if(ObjXml&&checkForm()){
		GetObject("Submit").value="提交信息...";
		GetObject("Submit").disabled=true;
		ObjXml.open("POST", Url+"Message.asp?Action=Save", true);
		ObjXml.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		var aIdArray=new Array("flag="+Math.random());
		var aUserArr=["SfMess_phone","SfMess_phone","SfMess_email","SfMess_content"];
		var argLen=aUserArr.length;
		for(i=0;i<argLen;i++){
			aIdArray[i+1]="&"+aUserArr[i]+"="+escape(GetObjValue(aUserArr[i]));
		}
//		this.getSex=function(){
//			var oSex=document.getElementsByName('Sex');
//			for(var i=0;i<oSex.length;i++){
//				if(oSex[i].checked){
//					return oSex[i].value;
//				}
//			}
//			return "M/F";
//		}
//		aIdArray[i+1]="&Sex="+this.getSex();
//		this.getShow=function(){
//			var Show=document.getElementsByName('Show');
//			for(var j=0;j<Show.length;i++){
//				if(Show[j].checked){
//					return Show[j].value;
//				}
//			}
//			return "0";
//		}
//		aIdArray[i+2]="&Show="+this.getShow();
		var data =aIdArray.join('');
	  //  alert(data);	
		ObjXml.onreadystatechange=function(){  //1
			if(ObjXml.readyState==4){ //2
				if(ObjXml.status==200){//3
					if(ObjXml.responseText==1){//4
						alert("留言成功");
					}else{//4
						alert("请您将主题/昵称/内容填写完整再提交");
					}//4

					GetObject("Submit").value="提交";
					GetObject("Submit").disabled=false;
					ClearForm();
					RefreshList();
				}else{//3
					alert("网络传输错误！请重试！");	
				}//3
			}//2	
		};//1
    		ObjXml.send(data);
  	}
} 

function ClearForm(){	//清空表单的函数
	GetObject("Title").value="";
	GetObject("NickName").value="";
	GetObject("QQ").value="";
	GetObject("Email").value="";
	GetObject("HomePage").value="http://17show.net";
	GetObject("Content").value="";
}

function RefreshList(){
	if(/LastDate=([^;]+)/.test(document.cookie)){
		var exp=new Date();
		exp.setTime(exp.getTime()-1);
		document.cookie="LastDate="+RegExp.$1+";expires="+exp.toGMTString();
	}
	GetList(1);
}

function MakePerNote(NickName,Sex,QQ,Email,HomePage,Title,Content,Reply,ReplyTime,DateAndTime,Ip){
    var onlineQQ,TaoTao;
    if(QQ=="N/F"){
       onlineQQ="<a target=\"blank\" href=\"tencent://message/?uin=25362779&Site=一起秀&Menu=yes\"><img border=\"0\" SRC=\"http://wpa.qq.com/pa?p=1:25362779:5\" alt=\"发送信息\"></a>";
       TaoTao="<a target=\"blank\" href=\"http://www.taotao.com/space.shtml?qq=25362779&invi=1\">滔滔</a>";
    }else{
       onlineQQ="<a target=\"blank\" href=\"tencent://message/?uin="+QQ+"&Site=一起秀&Menu=yes\"><img border=\"0\" SRC=\"http://wpa.qq.com/pa?p=1:"+QQ+":5\" alt=\"发送信息\"></a>";
       TaoTao="<a target=\"blank\" href=\"http://www.taotao.com/space.shtml?qq="+QQ+"&invi=1\">滔滔</a>";
    }
    if(Email=="N/F"){
       Email="<a href=\"mailto:25362779@qq.com\">Email</a>";
    }else{
       Email="<a href=\"mailto:"+Email+"\">Email</a>";
    }
    if(Reply!=="N/F"){
       Reply="<div class=\"NoteReply\"><b>站长回复：</b>"+unescape(Reply)+"["+ReplyTime+"]</div>"
    }else{
       Reply=""
    }
    HomePage="<a target=\"blank\" href=\""+unescape(HomePage)+"\">HomePage</a>"
	var tempStr='<div class="Note">\
	        <div class="NoteTopic"><b>主题：</b>'+unescape(Title)+'</div>\
			<div class="NoteMain">\
			     <div class="NoteVisualize"><div class="Sex">'+unescape(Sex)+'</div><div class="NickName">'+unescape(NickName)+'</div></div>\
                 <div class="NoteContent">'+unescape(Content)+Reply+'</div>\
            </div>\
            <div class="NotePerson">\
                 <div class="NoteTime">时间：'+DateAndTime+'</div>\
                 <div class="NoteElse">['+onlineQQ+']&nbsp;['+TaoTao+']&nbsp;['+Email+']&nbsp;['+HomePage+']&nbsp;['+Ip+']</div>\
            </div>\
		</div>';
	return tempStr;
}

function MakeNoteList(Str){	//输出服务器返回的留言内容
	if(Str!=0){
		var LeoBookList=eval("new Array("+Str+")");
		var allStr="";
		for(var i=0;i<LeoBookList.length;i++){
			allStr+=MakePerNote(LeoBookList[i].NickName,LeoBookList[i].Sex,LeoBookList[i].QQ,LeoBookList[i].Email,LeoBookList[i].HomePage,LeoBookList[i].Title,LeoBookList[i].Content,LeoBookList[i].Reply,LeoBookList[i].ReplyTime,LeoBookList[i].DateAndTime,LeoBookList[i].Ip);	
		}
	}else{	//返回0说明没有留言
		allStr="<div class=\"LeoTip\">暂时还没有留言！</div>"
	}
	GetObject("LeoBookList").innerHTML=allStr;
}

function GetList(page){	//获取指定页的留言
	GetObject("LeoBookList").innerHTML="<div class=\"LeoTip\">留言加载中....请稍后!</div>";	//清除原来显示的内容
	var ObjXml=ObjXML();
	ObjXml.open("GET", Url+"getRecord.asp?page="+page+"&Belong="+Sort+"&PageRecord="+PageRecord+"&r="+Math.random(), true);
	ObjXml.onreadystatechange=function(){
		if(ObjXml.readyState==4){
			if(ObjXml.status==200){
				MakeNoteList(ObjXml.responseText);
				GetPage();	//更新分页信息
			}else{
				alert("获取留言失败！请刷新重试！");	
			}
		}
		
	}
	ObjXml.send(null);
}

function UpdateList(){
	var ObjXml=ObjXML();
	ObjXml.open("GET", Url+"getRecord.asp?Action=GetUpdate&Belong="+Sort+"&PageRecord="+PageRecord+"&r="+Math.random(), true);
	ObjXml.onreadystatechange=function(){
		if(ObjXml.readyState==4){
			if(ObjXml.status==200){
				if(/LastDate=([^;]+)/.test(document.cookie) && unescape(RegExp.$1)!=ObjXml.responseText){
					GetList();
				}
				document.cookie="LastDate="+escape(ObjXml.responseText);
			}
		}
		
	}
	ObjXml.send(null);
	setTimeout("UpdateList()",15000);
}

function GetPage(){
	GetObject("DivPageList").innerHTML="<div id=\"loadPage\">分页信息加载中....请稍后!</div>";
	var ObjXml=ObjXML();
	var Leo=6;
	ObjXml.open("GET", Url+"getRecord.asp?Action=GetDivPageList&Belong="+Sort+"&PageRecord="+PageRecord+"&r="+Math.random(), true);
	ObjXml.onreadystatechange=function(){
		if(ObjXml.readyState==4){
			if(ObjXml.status==200){
				var Result=ObjXml.responseText.split("|");
				var CurrentPage=parseInt(Result[3],10);
				var AllPage=parseInt(Result[2],10);
				var tempPageStr=new Array("<div class=\"ForwardClass\"><div class=\"ForwardInformation\">共有"+Result[0]+"条留言&nbsp;第<b>"+CurrentPage+"</b>/<b>"+AllPage+"</b>页&nbsp;每页<b>"+Result[1]+"</b>条留言</div></div><div class=\"BackClass\">");
                var k;
                if(CurrentPage%Leo==0){
                   k=CurrentPage/Leo-1;
                }else{
                   k=(CurrentPage-CurrentPage%Leo)/Leo;
                } 

                if(Leo<CurrentPage){
                   tempPageStr[1]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage-Leo)+")\"><font face=\"webdings\">9</font></a></div>"
                }else{
                   tempPageStr[1]="<div class=\"RowClass\"><font face=\"webdings\">9</font></div>"
                }
                if(CurrentPage>1){
                   tempPageStr[2]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage-1)+")\"><font face=\"webdings\">9</font></a></div>"
                }else{
                   tempPageStr[2]="<div class=\"RowClass\"><font face=\"webdings\">7</font></div>"
                }
                
				for(var i=1;i<=Leo&(i+k-1)<AllPage;i++){
					if((k+i)!=CurrentPage)
						tempPageStr[i+2]="<div class=\"PageClass\"><a href=\"javascript:GetList("+(k+i)+")\">"+(k+i)+"</a></div>";
					else
						tempPageStr[i+2]="<div class=\"CurrentPageClass\">"+(k+i)+"</div>";
				}
                if(CurrentPage<AllPage){                           
                   tempPageStr[i+2]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage+1)+")\"><font face=\"webdings\" title=\"下一页\">8</font></a></div>"
                }else{
                   tempPageStr[i+2]="<div class=\"RowClass\"><font face=\"webdings\" title=\"下一页\">8</font></div>"
                }
                
                if(CurrentPage%Leo==0){                       
                   if(CurrentPage<AllPage) tempPageStr[i+3]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage+Leo)+")\"><font face=\"webdings\" title=\"下一页\">:</font></a></div>"
                }else if((CurrentPage-(CurrentPage%Leo)+Leo)<AllPage){
                   tempPageStr[i+3]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage+Leo)+")\"><font face=\"webdings\" title=\"下一页\">:</font></a></div>"
                }else{
                   tempPageStr[i+3]="<div class=\"RowClass\"><font face=\"webdings\" title=\"下一页\">:</font></div>"
                }
				tempPageStr[i+4]="<select name=\"page\" onchange=\"GetList(this.value)\" style=\"width:50px;\">";
				
				for(var j=1;j<=AllPage;j++){
					if(j!=CurrentPage)
						tempPageStr[i+4]=tempPageStr[i+4]+"<option value=\""+j+"\">"+j+"</option>";
					else
						tempPageStr[i+4]=tempPageStr[i+4]+"<option selected=\"selected\" value=\""+j+"\">"+j+"</option>";
				}
				tempPageStr[i+4]=tempPageStr[i+4]+"</select></div>";
				var ResultStr=tempPageStr.join('');
				GetObject("DivPageList").innerHTML="<div class=\"DividePageClass\">"+ResultStr+"</div>";
			}else{
				alert("获取分页信息失败！请刷新重试！");	
			}
		}
	}
	ObjXml.send(null);
}
function ChangeStyle(id){
	var stylesheet=GetObject("Color").href="Images/Leo"+id+".css";
	document.cookie="stylesheet="+escape(stylesheet);
}

function initStyle(){
		if(/stylesheet=([^;]+)/.test(document.cookie))
			GetObject("Color").href=unescape(RegExp.$1);
}
window.onload=function(){initStyle();GetList();UpdateList()}