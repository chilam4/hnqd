/*
������������Ӱ�����������Ӧ�ٶ�,�벻Ҫɾ����������!
�ٷ���ҳ:http://17show.net/book/
���򿪷�:������ţȥ����(Leo/Zero Crystal)
��Ȩ����:Copyright 2003-2005 һ��������(17show.net).INC All Right Reserved 
 */
var Url="";
var Sort="00000001";
var PageRecord=6;

function Trim(){
	return this.replace(/\s+$|^\s+/g,"");
}
String.prototype.Trim=Trim;	//�������˿ո�

function GetObject(elementId) { 	//��ȡָ��id��object
	if (document.getElementById) { 
		return document.getElementById(elementId); 
	} else if (document.all) { 
		return document.all[elementId]; 
	} else if (document.layers) { 
		return document.layers[elementId]; 
	} 
}

function GetObjValue(elementId){	//��ȡָ��id��form�����ֵ
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

function checkForm(){	//���ļ��
	if(GetObjValue("SfMess_content")==""||GetObjValue("SfMess_phone")==""||GetObjValue("SfMess_email")==""||GetObjValue("SfMess_phone")==""){
		alert("����������/�ǳ�/������д�������ύ��");
		return false;
	}
	if(GetObjValue("NickName").length>20){
		alert("�������Ҳ̫���˰ɣ���Ҫ����20Ŷ");
		return false;
	}
	if(GetObjValue("QQ")!="" && !/^[0-9]{5,10}$/.test(GetObjValue("qq"))){
		alert("��Ѷֻ��5-10λ��qq�Űɣ�");
		return false;
	}
	if(GetObjValue("Email")!="" && !/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/.test(GetObjValue("email"))){
		alert("�����ַ����");
		return false;
	}
	return true;
}

function ObjXmlSend() {	//��������
	var ObjXml=ObjXML();
	if(ObjXml&&checkForm()){
		GetObject("Submit").value="�ύ��Ϣ...";
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
						alert("���Գɹ�");
					}else{//4
						alert("����������/�ǳ�/������д�������ύ");
					}//4

					GetObject("Submit").value="�ύ";
					GetObject("Submit").disabled=false;
					ClearForm();
					RefreshList();
				}else{//3
					alert("���紫����������ԣ�");	
				}//3
			}//2	
		};//1
    		ObjXml.send(data);
  	}
} 

function ClearForm(){	//��ձ��ĺ���
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
       onlineQQ="<a target=\"blank\" href=\"tencent://message/?uin=25362779&Site=һ����&Menu=yes\"><img border=\"0\" SRC=\"http://wpa.qq.com/pa?p=1:25362779:5\" alt=\"������Ϣ\"></a>";
       TaoTao="<a target=\"blank\" href=\"http://www.taotao.com/space.shtml?qq=25362779&invi=1\">����</a>";
    }else{
       onlineQQ="<a target=\"blank\" href=\"tencent://message/?uin="+QQ+"&Site=һ����&Menu=yes\"><img border=\"0\" SRC=\"http://wpa.qq.com/pa?p=1:"+QQ+":5\" alt=\"������Ϣ\"></a>";
       TaoTao="<a target=\"blank\" href=\"http://www.taotao.com/space.shtml?qq="+QQ+"&invi=1\">����</a>";
    }
    if(Email=="N/F"){
       Email="<a href=\"mailto:25362779@qq.com\">Email</a>";
    }else{
       Email="<a href=\"mailto:"+Email+"\">Email</a>";
    }
    if(Reply!=="N/F"){
       Reply="<div class=\"NoteReply\"><b>վ���ظ���</b>"+unescape(Reply)+"["+ReplyTime+"]</div>"
    }else{
       Reply=""
    }
    HomePage="<a target=\"blank\" href=\""+unescape(HomePage)+"\">HomePage</a>"
	var tempStr='<div class="Note">\
	        <div class="NoteTopic"><b>���⣺</b>'+unescape(Title)+'</div>\
			<div class="NoteMain">\
			     <div class="NoteVisualize"><div class="Sex">'+unescape(Sex)+'</div><div class="NickName">'+unescape(NickName)+'</div></div>\
                 <div class="NoteContent">'+unescape(Content)+Reply+'</div>\
            </div>\
            <div class="NotePerson">\
                 <div class="NoteTime">ʱ�䣺'+DateAndTime+'</div>\
                 <div class="NoteElse">['+onlineQQ+']&nbsp;['+TaoTao+']&nbsp;['+Email+']&nbsp;['+HomePage+']&nbsp;['+Ip+']</div>\
            </div>\
		</div>';
	return tempStr;
}

function MakeNoteList(Str){	//������������ص���������
	if(Str!=0){
		var LeoBookList=eval("new Array("+Str+")");
		var allStr="";
		for(var i=0;i<LeoBookList.length;i++){
			allStr+=MakePerNote(LeoBookList[i].NickName,LeoBookList[i].Sex,LeoBookList[i].QQ,LeoBookList[i].Email,LeoBookList[i].HomePage,LeoBookList[i].Title,LeoBookList[i].Content,LeoBookList[i].Reply,LeoBookList[i].ReplyTime,LeoBookList[i].DateAndTime,LeoBookList[i].Ip);	
		}
	}else{	//����0˵��û������
		allStr="<div class=\"LeoTip\">��ʱ��û�����ԣ�</div>"
	}
	GetObject("LeoBookList").innerHTML=allStr;
}

function GetList(page){	//��ȡָ��ҳ������
	GetObject("LeoBookList").innerHTML="<div class=\"LeoTip\">���Լ�����....���Ժ�!</div>";	//���ԭ����ʾ������
	var ObjXml=ObjXML();
	ObjXml.open("GET", Url+"getRecord.asp?page="+page+"&Belong="+Sort+"&PageRecord="+PageRecord+"&r="+Math.random(), true);
	ObjXml.onreadystatechange=function(){
		if(ObjXml.readyState==4){
			if(ObjXml.status==200){
				MakeNoteList(ObjXml.responseText);
				GetPage();	//���·�ҳ��Ϣ
			}else{
				alert("��ȡ����ʧ�ܣ���ˢ�����ԣ�");	
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
	GetObject("DivPageList").innerHTML="<div id=\"loadPage\">��ҳ��Ϣ������....���Ժ�!</div>";
	var ObjXml=ObjXML();
	var Leo=6;
	ObjXml.open("GET", Url+"getRecord.asp?Action=GetDivPageList&Belong="+Sort+"&PageRecord="+PageRecord+"&r="+Math.random(), true);
	ObjXml.onreadystatechange=function(){
		if(ObjXml.readyState==4){
			if(ObjXml.status==200){
				var Result=ObjXml.responseText.split("|");
				var CurrentPage=parseInt(Result[3],10);
				var AllPage=parseInt(Result[2],10);
				var tempPageStr=new Array("<div class=\"ForwardClass\"><div class=\"ForwardInformation\">����"+Result[0]+"������&nbsp;��<b>"+CurrentPage+"</b>/<b>"+AllPage+"</b>ҳ&nbsp;ÿҳ<b>"+Result[1]+"</b>������</div></div><div class=\"BackClass\">");
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
                   tempPageStr[i+2]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage+1)+")\"><font face=\"webdings\" title=\"��һҳ\">8</font></a></div>"
                }else{
                   tempPageStr[i+2]="<div class=\"RowClass\"><font face=\"webdings\" title=\"��һҳ\">8</font></div>"
                }
                
                if(CurrentPage%Leo==0){                       
                   if(CurrentPage<AllPage) tempPageStr[i+3]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage+Leo)+")\"><font face=\"webdings\" title=\"��һҳ\">:</font></a></div>"
                }else if((CurrentPage-(CurrentPage%Leo)+Leo)<AllPage){
                   tempPageStr[i+3]="<div class=\"RowClass\"><a href=\"javascript:GetList("+(CurrentPage+Leo)+")\"><font face=\"webdings\" title=\"��һҳ\">:</font></a></div>"
                }else{
                   tempPageStr[i+3]="<div class=\"RowClass\"><font face=\"webdings\" title=\"��һҳ\">:</font></div>"
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
				alert("��ȡ��ҳ��Ϣʧ�ܣ���ˢ�����ԣ�");	
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