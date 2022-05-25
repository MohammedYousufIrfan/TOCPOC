
(function () {
    "use strict";
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            // Add a click event handler for the highlight button.
            $('#btnToc').click(AddToc);
        });
    };
    function AddToc() {
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;
            // Queue a commmand to get the HTML contents of the body.
            var bodyOOXML = body.getOoxml();
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                const url = "https://localhost:44339/TOC/addtoc";
                var data = { OOXML: bodyOOXML.value };
                $.ajax({
                    type: "POST",
                    url: url,
                    data: JSON.stringify(data),
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    async: false,
                    success: function (dat) {
                        setOOXML(dat.data);
                    }
                });
            });
        }).catch(errorHandler);
    }

    function setOOXML(currentOOXML) {
        // Check whether we have OOXML in the variable.
        if (currentOOXML != "") {
            // Run a batch operation against the Word object model.
            Word.run(function (context) {
                // Create a proxy object for the document body.
                var body = context.document.body;
                // Queue a commmand to insert OOXML in to the beginning of the body.
                body.insertOoxml(currentOOXML, Word.InsertLocation.replace);
                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync();
            }).catch(errorHandler);
        }
    }
    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
