
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                $('#italicize-button-text').text("Italicize!");
                $('#italicize-button-desc').text("Italicizes a word.");
                
                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("Please make use of the tools below to ease the process of manuscript writing.");

            // Text and description of the highlight button.

            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the longest word.");

            // Text and description of the italicize button.

            $('#italicize-button-text').text("Italicize Genes!");
            $('#italicize-button-desc').text("Italicizes all genes within the document.");

            // Text and description of the italicize button.

            $('#test-button-text').text("test");
            $('#test-button-desc').text("testing.");

            // Creates sample text to experiment with (currently disabled).
            //loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
            $('#italicize-button').click(italicizeGenes);
            $('#test-button').click(testFunction);

        });
    };

    //function loadSampleData() {
    //    // Run a batch operation against the Word object model.
    //    Word.run(function (context) {
    //        // Create a proxy object for the document body.
    //        var body = context.document.body;

    //        // Queue a commmand to clear the contents of the body.
    //        body.clear();
    //        // Queue a command to insert text into the end of the Word document body.
    //        body.insertText(
    //            "rhaT rhaK rhaS RhaK rpsD supE Mesopotamia Rhizobium United States of America",
    //            Word.InsertLocation.end);

    //        // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    //        return context.sync();
    //    })
    //    .catch(errorHandler);
    //}

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();
            
            // This variable will keep the search results for the longest word.
            var searchResults;
            
            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
        .catch(errorHandler);
    } 

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    function italicizeGenes() {
        Word.run(function (context) {
            OfficeExtension.config.extendedErrorLogging = true;

            // Search document for a word with lowercase, lowercase, lowercase, Uppercase.
            var searchResults = context.document.body.search('<[a-z][a-z][a-z][A-Z]>', {matchWildCards: true});

            // Load the search results.
            context.load(searchResults, 'font');

            return context.sync().then(function () {
                console.log("Gene count: " + searchResults.items.length);
                showNotification(searchResults.items.length + " gene names were italicized.");

                // Italicize each found gene.
                for (var i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].font.italic = true;
                }
                // Synchronize the document state.
                return context.sync();
            });
        })
        .catch(errorHandler);
    }

    function testFunction() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Queue a command to search the document with a wildcard
            // for any string of characters that starts with 'to' and ends with 'n'.
            var searchResults = context.document.body.search('to*n', { matchWildCards: true });

            // Queue a command to load the search results and get the font property values.
            context.load(searchResults, 'font');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Found count: ' + searchResults.items.length);

                // Queue a set of commands to change the font for each found item.
                for (var i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].font.color = 'purple';
                    searchResults.items[i].font.highlightColor = 'pink';
                    searchResults.items[i].font.bold = true;
                }

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync();
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
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
