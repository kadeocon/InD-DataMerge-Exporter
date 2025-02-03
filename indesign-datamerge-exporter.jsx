/*************************************************************************
 * InDesign Data Merge Exporter Script
 * Version: 1.0.0
 * Created: 2025-02-03
 * Author: Kade O'Connor
 * License: MIT
 * Repository: https://github.com/kadeocon/InD-DataMerge-Exporter
 *************************************************************************/

// Export InDesign pages as separate PDFs with custom naming
// @target indesign
// @targetengine "session"

$.strict = true;

// Polyfill indexOf for older InDesign versions
if (!Array.prototype.indexOf) {
    Array.prototype.indexOf = function(searchElement) {
        if (this === null) throw new TypeError();
        var n, k, t = Object(this), len = t.length >>> 0;
        if (len === 0) return -1;
        n = 0;
        if (arguments.length > 0) {
            n = Number(arguments[1]);
            if (n != n) n = 0;
            else if (n != 0 && n != Infinity && n != -Infinity) {
                n = (n > 0 || -1) * Math.floor(Math.abs(n));
            }
        }
        if (n >= len) return -1;
        for (k = n >= 0 ? n : Math.max(len - Math.abs(n), 0); k < len; k++) {
            if (k in t && t[k] === searchElement) return k;
        }
        return -1;
    };
}

// Function to clean text for filenames
function cleanTextForFilename(text) {
    // Handle empty or null text
    if (!text) return "unnamed";
    
    // First convert common typographic characters
    text = text.replace(/[—–]/g, "-");      // em and en dashes
    text = text.replace(/[""]/g, "");       // curly quotes
    text = text.replace(/['']/g, "");       // smart quotes
    text = text.replace(/[©®™]/g, "");      // copyright and trademark symbols
    text = text.replace(/[°]/g, "deg");     // degree symbol
    text = text.replace(/[&]/g, "and");     // ampersand
    text = text.replace(/[@]/g, "at");      // at symbol
    text = text.replace(/[$]/g, "");        // dollar sign
    text = text.replace(/[%]/g, "pct");     // percentage
    text = text.replace(/[#]/g, "num");     // hash/pound
    text = text.replace(/[+]/g, "plus");    // plus sign
    text = text.replace(/[=]/g, "eq");      // equals sign
    
    // Replace illegal filename characters and control characters
    text = text.replace(/[\/\\:*?"<>|]/g, "-");  // Windows/Unix illegal chars
    text = text.replace(/[\x00-\x1f\x7f]/g, ""); // Control characters
    
    // Convert all remaining non-alphanumeric characters to hyphens
    text = text.replace(/[^a-zA-Z0-9]/g, "-");
    
    // Replace multiple consecutive hyphens with a single hyphen
    text = text.replace(/-+/g, "-");
    
    // Remove hyphens from start and end
    text = text.replace(/^-+|-+$/g, "");
    
    // Ensure the text isn't empty after cleaning
    if (!text) return "unnamed";
    
    // Truncate if too long (common filesystem limit is 255 bytes)
    // Leave room for the prefix, suffix, and extension
    if (text.length > 200) {
        text = text.substr(0, 200);
        // Ensure we don't end with a partial word or hyphen
        text = text.replace(/-[^-]*$/, "");
    }
    
    return text;
}

// Initialize main function
var main = function() {
    if (!app.documents.length) {
        alert("No documents are open. Please open a document and try again.");
        return;
    }

    try {
        // Store original preferences to restore later
        var originalInteractionLevel = app.scriptPreferences.userInteractionLevel;
        var originalPreflightState = app.preflightOptions.preflightOff;

        // Disable user interaction and preflight temporarily
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
        app.preflightOptions.preflightOff = true;

        var doc = app.activeDocument;
        
        if (!doc.saved) {
            alert("Please save your document before running this script.");
            return;
        }

        // Get list of paragraph styles from document
        var styleNames = [];
        for (var i = 0; i < doc.paragraphStyles.length; i++) {
            styleNames.push(doc.paragraphStyles[i].name);
        }
        
        if (styleNames.length === 0) {
            alert("No paragraph styles found in document.");
            return;
        }

        // Create dialog
        var myDialog = new Window("dialog", "Export Settings");
        myDialog.orientation = "column";
        myDialog.alignChildren = "left";
        myDialog.preferredSize.width = 400;
        myDialog.spacing = 10;
        myDialog.margins = 16;

        // Export mode group
        var modeGroup = myDialog.add("panel", undefined, "Export Mode");
        modeGroup.orientation = "row";
        modeGroup.alignChildren = "left";
        modeGroup.margins = 15;
        var allPagesButton = modeGroup.add("radiobutton", undefined, "All Pages");
        var singlePageButton = modeGroup.add("radiobutton", undefined, "Single Page");
        allPagesButton.value = true;

        // Page number group
        var pageGroup = myDialog.add("group");
        pageGroup.orientation = "row";
        pageGroup.spacing = 10;
        pageGroup.add("statictext", undefined, "Page number:");
        var pageNumber = pageGroup.add("edittext", undefined, "1");
        pageNumber.characters = 5;
        pageNumber.enabled = false;

        // File prefix group
        var prefixGroup = myDialog.add("group");
        prefixGroup.orientation = "row";
        prefixGroup.spacing = 10;
        prefixGroup.add("statictext", undefined, "File prefix:");
        var prefix = prefixGroup.add("edittext", undefined, "2025-");
        prefix.characters = 40;

        // Paragraph style group
        var styleGroup = myDialog.add("group");
        styleGroup.orientation = "row";
        styleGroup.spacing = 10;
        styleGroup.add("statictext", undefined, "Paragraph style:");
        var styleDropdown = styleGroup.add("dropdownlist", undefined, styleNames);
        styleDropdown.selection = 0;
        styleDropdown.preferredSize.width = 300;

        // File suffix group
        var suffixGroup = myDialog.add("group");
        suffixGroup.orientation = "row";
        suffixGroup.spacing = 10;
        suffixGroup.add("statictext", undefined, "File suffix:");
        var suffix = suffixGroup.add("edittext", undefined, "-Market-Sheet");
        suffix.characters = 40;

        // File extension group
        var extGroup = myDialog.add("group");
        extGroup.orientation = "row";
        extGroup.spacing = 10;
        extGroup.add("statictext", undefined, "File extension:");
        var extension = extGroup.add("dropdownlist", undefined, ["pdf", "jpg", "png"]);
        extension.selection = 0;

        // Buttons group
        var buttonGroup = myDialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "center";
        buttonGroup.add("button", undefined, "OK", {name: "ok"});
        buttonGroup.add("button", undefined, "Cancel", {name: "cancel"});

        // Radio button event handlers
        allPagesButton.onClick = function() {
            pageNumber.enabled = false;
            pageNumber.text = "1";
        };

        singlePageButton.onClick = function() {
            pageNumber.enabled = true;
            pageNumber.active = true;
        };

        // Input validation
        pageNumber.onChanging = function() {
            if (this.text.match(/[^0-9]/)) {
                this.text = this.text.replace(/[^0-9]/g, "");
            }
        };

        if (myDialog.show() === 1) {
            // Get dialog values
            var exportSinglePage = singlePageButton.value;
            var selectedPage = parseInt(pageNumber.text, 10);
            var userPrefix = prefix.text;
            var userStyleName = styleDropdown.selection.text;
            var userSuffix = suffix.text;
            var userExtension = extension.selection.text;

            // Validate page number for single page export
            if (exportSinglePage) {
                if (isNaN(selectedPage) || selectedPage < 1 || selectedPage > doc.pages.length) {
                    alert("Invalid page number. Please enter a number between 1 and " + doc.pages.length);
                    return;
                }
            }

            // Set up file filter based on extension
            var fileFilter;
            switch(userExtension) {
                case "pdf":
                    fileFilter = "PDF Files:*.pdf";
                    break;
                case "jpg":
                    fileFilter = "JPEG Files:*.jpg,*.jpeg";
                    break;
                case "png":
                    fileFilter = "PNG Files:*.png";
                    break;
            }

            // Get initial filename for preview
            var previewName = "temp";
            if (exportSinglePage) {
                try {
                    var page = doc.pages[selectedPage - 1];
                    for (var j = 0; j < page.textFrames.length; j++) {
                        var textFrame = page.textFrames[j];
                        if (textFrame.paragraphs.length > 0 && 
                            textFrame.paragraphs[0].appliedParagraphStyle.name === userStyleName) {
                            previewName = cleanTextForFilename(textFrame.paragraphs[0].contents);
                            break;
                        }
                    }
                } catch (e) {
                    // If anything goes wrong, fall back to temp
                    previewName = "temp";
                }
            }

            // Get export location using InDesign's native dialog
            var defaultFolder = doc.filePath || Folder.desktop;
            var initialFilename = userPrefix + previewName + userSuffix + "." + userExtension;
            var tempFile = File(defaultFolder + "/" + initialFilename);
            var exportPath = tempFile.saveDlg("Choose export location", fileFilter);

            if (!exportPath) return;
            var exportFolder = exportPath.parent;

            // Validate export folder
            if (!exportFolder.exists) {
                if (confirm("Folder does not exist. Create it?")) {
                    try {
                        if (!exportFolder.create()) {
                            throw new Error("Could not create folder");
                        }
                    } catch (e) {
                        if (!confirm("Could not create folder. Continue anyway?")) {
                            return;
                        }
                    }
                } else {
                    return;
                }
            }

            // Test write permissions
            var canWrite = true;
            try {
                var testFile = File(exportFolder + "/.test");
                testFile.open('w');
                testFile.close();
                testFile.remove();
            } catch (e) {
                canWrite = false;
            }

            if (!canWrite) {
                if (!confirm("Could not verify write permissions. This might be normal for cloud storage. Continue anyway?")) {
                    return;
                }
            }

            // Show progress window
            var progressWin = new Window("palette", "Exporting...");
            progressWin.progressBar = progressWin.add("progressbar", undefined, 0, 100);
            progressWin.progressBar.preferredSize.width = 300;
            progressWin.show();

            try {
                // Process pages
                var pagesToProcess = exportSinglePage ? [doc.pages[selectedPage - 1]] : doc.pages;

                for (var i = 0; i < pagesToProcess.length; i++) {
                    var page = pagesToProcess[i];
                    var pageIndex = exportSinglePage ? selectedPage - 1 : i;

                    // Update progress
                    progressWin.progressBar.value = (i / pagesToProcess.length) * 100;

                    // Find text with specified style
                    var identifierText = "";
                    var textFrames = page.textFrames;
                    var foundStyle = false;

                    for (var j = 0; j < textFrames.length; j++) {
                        var textFrame = textFrames[j];
                        if (textFrame.paragraphs.length > 0 && 
                            textFrame.paragraphs[0].appliedParagraphStyle.name === userStyleName) {
                            identifierText = textFrame.paragraphs[0].contents;
                            identifierText = cleanTextForFilename(identifierText);
                            foundStyle = true;
                            break;
                        }
                    }

                    if (!foundStyle) {
                        alert("Could not find text with paragraph style '" + userStyleName + 
                              "' on page " + (pageIndex + 1));
                        continue;
                    }

                    // Export file
                    var filename = userPrefix + identifierText + userSuffix + "." + userExtension;
                    var filePath = File(exportFolder + "/" + filename);

                    try {
                        switch(userExtension) {
                            case "pdf":
                                app.pdfExportPreferences.pageRange = (pageIndex + 1).toString();
                                doc.exportFile(ExportFormat.PDF_TYPE, filePath, false);
                                break;
                            case "jpg":
                                app.jpegExportPreferences.pageRange = (pageIndex + 1).toString();
                                doc.exportFile(ExportFormat.JPG, filePath, false);
                                break;
                            case "png":
                                app.pngExportPreferences.pageRange = (pageIndex + 1).toString();
                                doc.exportFile(ExportFormat.PNG_FORMAT, filePath, false);
                                break;
                        }
                    } catch (e) {
                        alert("Error exporting page " + (pageIndex + 1) + ": " + e.message);
                    }
                }

                progressWin.close();
                alert("Export complete! Files saved to:\n" + exportFolder.fsName);

            } catch (e) {
                alert("An error occurred during export: " + e.message);
            } finally {
                if (progressWin && progressWin.visible) {
                    progressWin.close();
                }
            }
        }

    } catch (e) {
        alert("An error occurred: " + e.message);
    } finally {
        // Restore original preferences
        app.scriptPreferences.userInteractionLevel = originalInteractionLevel;
        app.preflightOptions.preflightOff = originalPreflightState;
    }
};

// Execute main function
try {
    main();
} catch (e) {
    alert("Fatal Error: " + e.message);
}
