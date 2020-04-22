// JavaScript Document
var XMLHttp = {
    _objPool: [],
    
    _getInstance: function ()
    {
        for (var i = 0; i < this._objPool.length; i ++)
        {
            if (this._objPool[i].readyState == 0 || this._objPool[i].readyState == 4)
            {
                return this._objPool[i];
            }
        }
        
        this._objPool[this._objPool.length] = this._createObj();

        return this._objPool[this._objPool.length - 1];
    },

    _createObj: function ()
    {
        if (window.XMLHttpRequest)
        {
            var objXMLHttp = new XMLHttpRequest();

        }
        else
        {
            var MSXML = ['MSXML2.XMLHTTP.5.0', 'MSXML2.XMLHTTP.4.0', 'MSXML2.XMLHTTP.3.0', 'MSXML2.XMLHTTP', 'Microsoft.XMLHTTP'];
            for(var n = 0; n < MSXML.length; n ++)
            {
                try
                {
                    var objXMLHttp = new ActiveXObject(MSXML[n]);        
                    break;
                }
                catch(e)
                {
                }
            }
         }

        if (objXMLHttp.readyState == null)
        {
            objXMLHttp.readyState = 0;

            objXMLHttp.addEventListener("load", function ()
                {
                    objXMLHttp.readyState = 4;
                    
                    if (typeof objXMLHttp.onreadystatechange == "function")
                    {
                        objXMLHttp.onreadystatechange();
                    }
                },  false);
        }

        return objXMLHttp;
    },
    
    sendReq: function (method, url, data, callback, callValue)
    {
        var objXMLHttp = this._getInstance();

        with(objXMLHttp)
        {
            try
            {
                if (url.indexOf("?") > 0)
                {
                    url += "&randnum=" + Math.random();
                }
                else
                {
                    url += "?randnum=" + Math.random();
                }

                open(method, url, true);
                setRequestHeader('Content-Type', 'application/x-www-form-urlencoded; charset=gb2312');
                send(data);
                onreadystatechange = function ()
                {
                    if (objXMLHttp.readyState == 4)
                    {
                        callback(objXMLHttp, callValue);
                    }
                }
            }
            catch(e)
            {
                alert(e);
            }
        }
    }
};

function xmlOnLoad(obj, val)
{
	if(200 == obj.status || 304 == obj.status)
	{
		var tStr = obj.responseText;
		if(val)
		{
			var tO = eval('document.getElementById(\'' + val + '\')');
			if(tO){
				tO.innerHTML = tStr;
			}else{
				eval(val + '(\'' + tStr + '\')');
			}
		}else{
			eval(tStr);
		}
	}
}


function toquery(url){
var objxml=new ActiveXObject("Microsoft.XMLHTTP")
objxml.open("GET",url,false);
objxml.send();
if (objxml.status=="200")
{return objxml.responseText;}
else{return "err";}
}