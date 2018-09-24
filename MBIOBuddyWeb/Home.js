
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
                $("#tool-description").text("This sample displays the selected text.");
                $('#latin-button-text').text("Display!");
                $('#latin-button-desc').text("Display the selected text");
                $('#gene-button-text').text("Italicize gene!");
                $('#gene-button-desc').text("Italicizes a word.");
                $('#Tn-button-text').text("Italicize Tn!");
                $('#Tn-button-desc').text("Italicizes all transposon numbers within the document.");
                
                $('#latin-button').click(displaySelectedText);
                $('#italicize-button').click(displaySelectedText);
                $('#Tn-button').click(displaySelectedText);
                return;
            }

            $("#tool-description").text("Please make use of the tools below to ease the process of manuscript writing. I hope you find it useful.");

            // Text and description of the highlight button.

            $('#latin-button-text').text("Italicize Latin!");
            $('#latin-button-desc').text("Italicize common Latin abbreviations.");

            // Text and description of the italicize genes button.

            $('#gene-button-text').text("Italicize Genes!");
            $('#gene-button-desc').text("Italicizes all genes within the document.");

            // Text and description of the italicize transposons button.

            $('#Tn-button-text').text("Italicize Tn!");
            $('#Tn-button-desc').text("Italicizes all transposon numbers within the document.");

            // Creates sample text to experiment with (currently disabled).
            //loadSampleData();

            // Add a click event handler for the highlight button.
            $('#latin-button').click(italicizeLatin);
            $('#gene-button').click(italicizeGenes);
            $('#Tn-button').click(italicizeTn);

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
    //            "rhaT rhaK rhaS RhaK rpsD38 supE rhaSTPQ RhaSTPQ rhaDI RhaDI Mesopotamia Rhizobium United States of America Tn3 Tn5 Tn7 Tn10 10Tn10 10",
    //            Word.InsertLocation.end);

    //        // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    //        return context.sync();
    //    })
    //    .catch(errorHandler);
    //}

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

    function italicize(search) {
        Word.run(function (context) {
            OfficeExtension.config.extendedErrorLogging = true;

            if (search.items.length > 0) {
                for (var index = 0; index < search.items.length; index++) {
                    search.items[index].font.italic = true;
                }
            }
            return context.sync();
        })
        .catch(errorHandler);
    }

    function italicizeLatin() {
        Word.run(function (context) {
            OfficeExtension.config.extendedErrorLogging = true;

            // Search document for the most common Latin terms.
            var deNovo = context.document.body.search('de novo', { matchWildCards: true });
            var etAl = context.document.body.search('et al.', { matchWildCards: true });
            var inSitu = context.document.body.search('in situ', { matchWildCards: true });
            var inTrans = context.document.body.search('in trans', { matchWildCards: true });
            var inVitro = context.document.body.search('in vitro', { matchWildCards: true });
            var inVivo = context.document.body.search('in vivo', { matchWildCards: true });

            // Load the search results.
            context.load(deNovo, 'font');
            context.load(etAl, 'font');
            context.load(inSitu, 'font');
            context.load(inTrans, 'font');
            context.load(inVitro, 'font');
            context.load(inVivo, 'font');

            return context.sync().then(function () {
                var totalAmount = deNovo.items.length + etAl.items.length + inSitu.items.length + inTrans.items.length + inVitro.items.length + inVivo.items.length;
                console.log("Latin expressions count: " + totalAmount);
                showNotification(totalAmount + " Latin terms were italicized.");

                // Italicize every result.
                italicize(deNovo);
                italicize(etAl);
                italicize(inSitu);
                italicize(inTrans);
                italicize(inVitro);
                italicize(inVivo);

                return context.sync();
            });
        })
        .catch(errorHandler);
    }

    function italicizeGenes() {
        Word.run(function (context) {
            OfficeExtension.config.extendedErrorLogging = true;

            // Search document for a word with lowercase, lowercase, lowercase, Uppercase.
            var geneSearch = context.document.body.search('<[a-z][a-z][a-z][A-Z]', { matchWildCards: true });
            var clusterSearch = context.document.body.search('<[a-z][a-z][a-z][A-Z]@>', { matchWildCards: true });

            // Load the search results.
            context.load(geneSearch, 'font');
            context.load(clusterSearch, 'font');

            return context.sync().then(function () {
                console.log("Gene count: " + geneSearch.items.length);
                console.log("Gene cluster count: " + clusterSearch.items.length);
                showNotification(geneSearch.items.length + " gene names were italicized.");

                // Italicize each found gene.
                italicize(geneSearch);
                italicize(clusterSearch);

                // Synchronize the document state.
                return context.sync();
            });
        })
        .catch(errorHandler);
    }

    // Currently unoptimized. Hacked together a method that technically works, but is very memory inefficient. Thankfully, these are Word documents.
    function italicizeTn() {
        Word.run(function (context) {
            OfficeExtension.config.extendedErrorLogging = true;

            // Search document for a word with a word that begins with Tn and has number (not Tn10).
            var TnSearch = context.document.body.search('<Tn[0-9]', { matchWildCards: true });

            // Searches for Tn to un-do the italicization on the word.
            var Tns = context.document.body.search('Tn', { matchPrefix: true });

            // Searches for Tn10. I couldn't find any other way to do it.
            var Tn10 = context.document.body.search('Tn10', { matchPrefix: true });

            // Load the search results.
            context.load(TnSearch, 'font');
            context.load(Tns, 'font');
            context.load(Tn10, 'font');

            return context.sync().then(function () {
                console.log("Transposon count: " + TnSearch.items.length);
                console.log("Tn10 count: " + Tn10.items.length);
                showNotification(TnSearch.items.length + " transposon numbers were italicized.");

                // Italicize the third character of each transposon. This works because Tn5 
                for (var i = 0; i < TnSearch.items.length; i++) {
                    
                    // If the Tn number is over 9 (ie. Tn10), italicize both numbers. Currently not working.
                    //if (isNumeric(searchResults.items[i][3])) {
                    //    searchResults.items[i].font.italic = true;
                    
                    TnSearch.items[i].font.italic = true;
                    Tns.items[i].font.italic = false;
                           
                }

                if (Tn10.items.length > 0) {
                    for (var j = 0; j < TnSearch.items.length; j++) {
                        if (j < Tn10.items.length) {
                            Tn10.items[j].font.italic = true;
                        }
                        Tns.items[j].font.italic = false;
                    }
                }
                // Synchronize the document state.
                return context.sync();
            });
        })
            .catch(errorHandler);
    }

    // An accessory function to determine whether a character is numeric. Similar to JQuery's isNumeric() but in pure JS.
    function isNumeric(n) {
        return !isNaN(parseFloat(n)) && isFinite(n);
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
