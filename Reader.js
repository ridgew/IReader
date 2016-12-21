/********************************************************
iBTree for UL/LI List
Author: Ridge Wong 2005Äê9ÔÂ29ÈÕ 16:09:22
Version: 1.0 (First Edition)
Requirement:DOM 1.0 & Javascript Enabled Browser


****************Example***************
step 1 : Needed Style
-------------
<style type="text/css">
li {list-style-type:none;}
.folder {display:none;}
.box {cursor:pointer;}
</style>

step 2 : Needed Pics
---------------------
plus.gif + minus.gif
folder.gif + open.gif

step 3: Html Code demo
-------------------------------
<ul>
<li> <span class="box" onclick="iBTree(this,this.parentNode);"><img src="images/plus.gif" border="0"> <img src="images/folder.gif" border="0"></span> Chapter one
	<ul class="folder">
		<li> topic one
		<li> topic two
		<li> topic three
	</ul>
<li> Chapter Two
<li> Chapter Three
</ul>
*********************************************************/

function iBTree(caller,obj) {
	var k = obj.childNodes.length;
	for (var i=0; i<k ;i++ )
	{
		var _obj = obj.childNodes.item(i);
		if (_obj.childNodes.length>0)
		{
			//alert("TagName " +_obj.tagName + " ; with Child Length " +_obj.childNodes.length);
			if (_obj.tagName == "UL")
			{
				_obj.style.display = (_obj.style.display=="block") ? "none" : "block";
			}
		}
	}
	var strHtml = caller.innerHTML;
	strHtml = (strHtml.indexOf("plus.gif")!=-1) ? strHtml.replace(/plus\.gif/i,"minus.gif") : strHtml.replace(/minus\.gif/i,"plus.gif");

    /* --------- fixed folder ---------- */
	strHtml = (strHtml.indexOf("folder.gif")!=-1) ? strHtml.replace(/folder\.gif/i,"open.gif") : strHtml.replace(/open\.gif/i,"folder.gif");
	/*----------------------------------------*/
	caller.innerHTML = strHtml;
	//alert(strHtml);
}