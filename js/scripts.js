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
    var subject = Office.context.mailbox.item.subject;

    Office.context.mailbox.item.notificationMessages.addAsync("subject", {
        type: "informationalMessage",
        icon: "iconid",
        message: "Subject: " + subject,
        persistent: false
    });

    event.completed();
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
