/* taskpane.js */

var _om;
var _item;

Office.initialize = function (reason)
{
	_om = Office.context.mailbox;
	_item = _om.item;
}

function getSubject()
{
	document.getElementById("subject").innerHTML = _item.subject
}