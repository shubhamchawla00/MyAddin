debugger;
importScripts('office.debug.js');

self.addEventListener('message', function (e) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "succeeded") {
        }
    });
}, false);