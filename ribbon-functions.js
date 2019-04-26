(function () {
    Office.initialize = async function (reason) {
    };
})();

function openDialog(event) {
    var dialogUrl = "https://static-resources.yasoon.com/owa-display-dialog-issue/dialog.html";
    Office.context.ui.displayDialogAsync(dialogUrl, { height: 50, width: 50, displayInIframe: false }, function (asyncResult) {
        var dialog = asyncResult.value;

        // Handle close event
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            event.completed();
        });
    });
}