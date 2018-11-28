/* taskpane.js */

var _om;
var _item;

Office.initialize = function (reason)
{
	_om = Office.context.mailbox;
	_item = _om.item;
};

function getSubject()
{
	document.getElementById("subject").innerHTML = _item.subject
}

function uilessGetSubject(event)
{
Office.context.mailbox.item.body.setAsync(
 'Set body using UILess', 
 {coercionType: "text"}, 
 function (asyncResult) {
  if (asyncResult.status == "failed") {
	event.completed();
   	console.log("Action failed with error: " + asyncResult.error.message);
  } else {
	event.completed();
   	console.log("Successfully set body text");
  }
 }
);
}

function uilessAddNotification(event)
{
	_item.notificationMessages.addAsync("progress", { 
		type: "progressIndicator", 
		message : "Handling popups"
	});

	setTimeout(function () {
		Office.context.ui.displayDialogAsync("https://bing.com", {width: 50, height: 50}, function (asyncResult) {
			if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
				showMessage("Failed to show dialog: " + asyncResult.error.message);
				return;
			}
		
			_item.notificationMessages.removeAsync("progress");
		
			_item.notificationMessages.addAsync("information", {
				type: "informationalMessage",
				message: "Result: " + asyncResult.status,
				icon: "iconid",
				persistent: false
			});

			event.completed(true);
		})
	}, 10000);
}
