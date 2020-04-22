<!--
//	检查咨询反馈
//////////////////////////////////////////////////////////////////////////////////////////
function CheckFeedback(FormName)
{
	//判断输入的内容
	if ( IsBlank( FormName.Classid,"咨询类型" ) )
		return false;
	if ( IsBlank( FormName.Author,"您的姓名" ) )
		return false;
	if ( IsBlank( FormName.Email,"电子邮件" ) )
		return false;
	if ( CheckEmail( FormName.Email,false ) ) 
		return false;
	if ( IsBlank( FormName.FeedbackTitle,"咨询主题" ) )
		return false;
	if ( IsBlank( FormName.FeedbackContent,"咨询内容" ) )
		return false;
	FormName.submit();
}
//////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////
//判断输入的内容是否为空
//	var obj_form = document.FormMailto;
//	if ( IsBlank( obj_form.user_name,"用户帐号" ) ) 
//		return false;
//////////////////////////////////////////////////////////////////////////////////////////
function IsBlank( ObjInput, AlertTxt )
{
  if(ObjInput.value == "")
  {
	  alert(' 系统提示：\n\n'+AlertTxt+'不能为空，请您输入 ！\n');
      ObjInput.focus();
	  return true;
   }
}
//////////////////////////////////////////////////////////////////////////////////////////
//判断限制输入内容的字数
//<textarea name="GuestContent" cols="54" rows="10"  class="textarea1" onFocus="this.select()" onMouseOver="this.style.background='#FFFFFF';this.focus()" onMouseOut="this.style.background='#F7F7F7';" onkeydown=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain); onkeyup=gbcount(this.form.GuestContent,this.form.total,this.form.used,this.form.remain);></textarea>
//<br>最多字数：<input name=total disabled class="input" value=500 size=3 maxLength=4>已用字数：<input name=used disabled class="input" value=0 size=3 maxLength=4>剩余字数：<input name=remain disabled class="input" value=500 size=3 maxLength=4>
//////////////////////////////////////////////////////////////////////////////////////////
function gbcount(message,total,used,remain)
{
	var max;
	max = total.value;
	if (message.value.length > max) {
	message.value = message.value.substring(0,max);
	used.value = max;
	remain.value = 0;
	alert(' 系统提示：\n\n内容字数不能超过 1000 个字 ！\n');
	
		}
	else {
	used.value = message.value.length;
	remain.value = max - used.value;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////
// 修改编辑栏高度
//////////////////////////////////////////////////////////////////////////////////////////
function admin_Size(num,objname)
{
	var obj=document.getElementById(objname)
	if (parseInt(obj.rows)+num>=3) {
		obj.rows = parseInt(obj.rows) + num;	
	}
	if (num>0)
	{
		obj.width="90%";
	}
}

//////////////////////////////////////////////////////////////////////////////////////////
//判断输入的电子邮件
//	var ObjFormMailto = document.FormMailto;
//	if ( CheckEmail( ObjFormMailto.UserEmail,false ) ) 
//		return false;
//////////////////////////////////////////////////////////////////////////////////////////
function CheckEmail(ObjInput, AllowedNull)
{
	var datastr = ObjInput.value;
	var lefttrim = datastr.search(/\S/gi);
	
	if (lefttrim == -1) 
	{
		if (AllowedNull) 
		{
			return 1;
		} 
		else 
		{
			alert(" 系统提示：\n\n请您输入一个正确的E-mail地址！\n");
			ObjInput.focus();
			return -1;
		}
	}
	
	var myRegExp = /[a-z0-9](([a-z0-9]|[_\-\.][a-z0-9])*)@([a-z0-9]([a-z0-9]|[_\-][a-z0-9])*)((\.[a-z0-9]([a-z0-9]|[_\-][a-z0-9])*)*)/gi;
	var answerind = datastr.search(myRegExp);
	var answerarr = datastr.match(myRegExp);
	
	if (answerind == 0 && answerarr[0].length == datastr.length)
	{
		return 0;
	}

	alert(" 系统提示：\n\n请您输入一个正确的E-mail地址！\n");
	ObjInput.focus();
	return -1;
}


//-->
