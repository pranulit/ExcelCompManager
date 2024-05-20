// Create a new floating palette window
var myWindow = new Window("palette", "Comp-Sheet-Flow", undefined);
myWindow.orientation = "column";

// Button for 'Comps from Names'
var button1 = myWindow.add("button", undefined, "Comps from Names");
button1.onClick = function() {
    compsFromNames();
}

// Button for 'Content Extract'
var button2 = myWindow.add("button", undefined, "Content Extract");
button2.onClick = function() {
    contentExtract();
}

// Button for 'Content Import'
var button3 = myWindow.add("button", undefined, "Content Import");
button3.onClick = function() {
    contentImport();
}

// Display the window
myWindow.center();
myWindow.show();

// Function definitions
function compsFromNames() {
    var myPanel = (this instanceof Panel) ? this : new Window("palette", "Create Compositions", undefined, {resizeable:true});

    if (myPanel != null) {
        var textGroup = myPanel.add("group");
        textGroup.orientation = "row";
        textGroup.add("statictext", undefined, "Enter Composition Names:");

        var inputGroup = myPanel.add("group");
        inputGroup.orientation = "row";
        var compsInput = inputGroup.add("edittext", undefined, "", {multiline: true});
        compsInput.size = [300, 100];

        var fpsGroup = myPanel.add("group");
        fpsGroup.orientation = "row";
        fpsGroup.add("statictext", undefined, "Select FPS:");
        var fpsDropdown = fpsGroup.add("dropdownlist", undefined, ["24", "25", "30", "60"]);
        fpsDropdown.selection = 1;

        var buttonGroup = myPanel.add("group");
        buttonGroup.orientation = "row";
        var createBtn = buttonGroup.add("button", undefined, "Create Compositions");
        createBtn.onClick = function() {
            createCompositions(compsInput.text, parseFloat(fpsDropdown.selection.text));
        }
    }

    function createCompositions(compsInput, fps) {
        var lines = compsInput.split("\n");
        var mainFolder = app.project.items.addFolder("main"); // Create the main folder

        var compSettings = {
            "16x9": {width: 1920, height: 1080},
            "1x1": {width: 1080, height: 1080},
            "4x5": {width: 1080, height: 1350},
            "9x16": {width: 1080, height: 1920}
        };
        var defaultFormat = "16x9";

        app.beginUndoGroup("Create Compositions");

        try {
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i].replace(/^\s+|\s+$/g, ''); // Trim whitespace using regex
                if (line !== "") {
                    var parts = line.split("_");
                    var duration = parts[2].match(/(\d+)s/);
                    var formatMatch = parts[2].match(/(\d+x\d+)/);
                    var format = formatMatch ? formatMatch[1] : defaultFormat;

                    if (duration && compSettings[format]) {
                        var comp = app.project.items.addComp(parts[0], compSettings[format].width, compSettings[format].height, 1, parseFloat(duration[1]), fps);
                        comp.parentFolder = mainFolder; // Assign composition to the main folder
                        comp.name = line;

                        var solid = comp.layers.addSolid([1, 1, 1], ">offline", comp.width, comp.height, 1);
                        solid.name = ">offline";
                    } else {
                        alert("Could not determine valid duration/format for: " + line);
                    }
                }
            }
        } catch (e) {
            alert("Error: " + e.toString());
        }

        app.endUndoGroup();
    }

    if (myPanel instanceof Window) {
        myPanel.center();
        myPanel.show();
    } else {
        myPanel.layout.layout(true);
    }
}

function contentExtract() {
    // Define a function to find the index of an element in an array
    function indexOf(array, element) {
        // Loop through the array
        for (var i = 0; i < array.length; i++) {
        // If the current element matches the search element, return its index
        if (array[i] === element) {
            return i;
        }
        }
        // Return -1 if the element is not found in the array
        return -1;
    }
    
    // Define a function to create a user interface window in After Effects
    function createUI() {
        // Create a new dialog window titled "Select Compositions"
        var dlg = new Window("dialog", "Select Compositions", undefined, {
        resizeable: false,
        });
        dlg.size = [400, 400];
    
        // Add a group for the search bar
        var searchGroup = dlg.add("group");
        searchGroup.orientation = "row";
        // Add a text edit field and a search button to the group
        var searchEdit = searchGroup.add("edittext", undefined, "");
        searchEdit.characters = 20;
        var searchButton = searchGroup.add("button", undefined, "Search");
    
        // Add a list box to display items with multi-selection enabled
        var listBox = dlg.add("listbox", [0, 0, 380, 300], [], { multiselect: true });
        listBox.alignment = "fill";
    
        // Populate the list box initially
        populateListBox(listBox, ""); // Assuming a function populateListBox is defined
    
        // Define the search button click behavior
        searchButton.onClick = function () {
        populateListBox(listBox, searchEdit.text);
        };
    
        // Add a group for buttons
        var buttons = dlg.add("group");
        buttons.alignment = "right";
        var cancelButton = buttons.add("button", undefined, "Cancel", {
        name: "cancel",
        });
        var saveButton = buttons.add("button", undefined, "Save", { name: "ok" });
    
        // Define the cancel button behavior
        cancelButton.onClick = function () {
        dlg.close();
        };
        // Define the save button behavior
        saveButton.onClick = function () {
        try {
            var selectedCompositions = getSelectedCompositions(listBox);
            alert("Selected " + selectedCompositions.length + " compositions.");
            if (selectedCompositions.length === 0) {
            alert("No compositions are selected or associated properly.");
            return; // Exit if no compositions to process
            }
            exportSelectedCompositions(selectedCompositions);
            dlg.close();
        } catch (error) {
            alert("Error: " + error.toString());
        }
        };
    
        // Lay out the dialog elements, center it, and display it
        dlg.layout.layout(true);
        dlg.center();
        dlg.show();
    }
    
    // Define a function to get the compositions selected in the list box
    function getSelectedCompositions(listBox) {
        var selectedItems = [];
        // Loop through the list box items
        for (var i = 0; i < listBox.items.length; i++) {
        // Check if the item is selected
        if (listBox.items[i].selected) {
            // Check if associated composition exists to avoid errors
            if (listBox.items[i].compItem) {
            selectedItems.push(listBox.items[i].compItem);
            } else {
            alert("Error: No associated composition for selected item.");
            }
        }
        }
        return selectedItems; // Return the array of selected compositions
    }
    
    // Define a function to populate the list box with compositions that match the search term
    function populateListBox(listBox, searchTerm) {
        listBox.removeAll(); // Remove all current items from the list box
        var project = app.project;
        var compositions = project.items;
        // Loop through the compositions in the project
        for (var i = 1; i <= compositions.length; i++) {
        var compItem = compositions[i];
        // Check if the item is a composition and matches the search term
        if (compItem.name.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1) {
            var item = listBox.add("item", compItem.name);
            item.compItem = compItem;
        }
        }
    }
    
    function searchPrecomps(comp, parentCompData, project) {
        if (!parentCompData.textLayers) {
        parentCompData.textLayers = {};
        }
        if (!parentCompData.fileLayers) {
        parentCompData.fileLayers = {};
        }
        if (!parentCompData.specialLayers) {
        parentCompData.specialLayers = {};
        }
    
        for (var k = 1; k <= comp.numLayers; k++) {
        var layer = comp.layer(k);
    
        // Recursively handle precompositions
        if (layer.source instanceof CompItem) {
            searchPrecomps(layer.source, parentCompData, project);
        }
    
        // Extract text from text layers
        if (layer instanceof TextLayer && layer.property("Source Text") != null) {
            var textSource = layer.property("Source Text").value;
            if (textSource) {
            var textContent = textSource.text.replace(/[\r\n]+/g, " ");
            parentCompData.textLayers[layer.name] = textContent;
            }
        }
    
        // Capture all layers that start with ">", even if they don't have a file
        if (layer.name.substring(0, 1) === ">") {
            // If the layer has a source file, store the path; otherwise, note the absence of a path
            var filePath = (layer.source && layer.source.file) ? layer.source.file.fsName : "no path retrieved";
            parentCompData.fileLayers[layer.name] = filePath;
        }
    
        // Check if layer name starts with "#"
        if (layer.name.charAt(0) === "#") {
            // Extract and store the source name if the layer has a source
            if (layer.source) {
            var sourceName = getSourceName(layer.source.name, project);
            if (sourceName) {
                parentCompData.specialLayers[layer.name] = sourceName;
            } else {
                // Handle cases where source name is not found
                parentCompData.specialLayers[layer.name] = "Source not found";
            }
            } else {
            // Handle cases where there is no source
            parentCompData.specialLayers[layer.name] = "No source associated";
            }
        }
        }
    }
    
    // Helper function to get source name from project
    function getSourceName(layerName, project) {
        for (var i = 1; i <= project.items.length; i++) {
        var item = project.items[i];
        if (item instanceof FootageItem && item.name === layerName) {
            return item.name;
        }
        }
        return null; // Return null if source name is not found
    }
    
    function exportSelectedCompositions(compCheckboxes) {
        var comps = [];
        var uniqueTextLayerNames = {};
        var uniqueFileLayerNames = {};
        var uniqueSpecialLayerNames = {}; // Initialize for special layers
    
        // Get the After Effects project object
        var project = app.project;
    
        for (var i = 0; i < compCheckboxes.length; i++) {
        var compItem = compCheckboxes[i];
        if (!compItem || !compItem.name || !compItem.numLayers) {
            continue; // Skip if compItem is invalid
        }
    
        var parentCompData = { textLayers: {}, fileLayers: {}, specialLayers: {} };
        searchPrecomps(compItem, parentCompData, project); // Pass the project object
    
        var compObj = {
            name: compItem.name,
            textLayersContent: parentCompData.textLayers,
            fileLayersContent: parentCompData.fileLayers,
            specialLayersContent: parentCompData.specialLayers, // Add special layers content
        };
    
        comps.push(compObj);
    
        // Update unique layer names
        for (var layerName in parentCompData.textLayers) {
            uniqueTextLayerNames[layerName] = true;
        }
        for (var layerName in parentCompData.fileLayers) {
            uniqueFileLayerNames[layerName] = true;
        }
        for (var layerName in parentCompData.specialLayers) {
            uniqueSpecialLayerNames[layerName] = true; // Collect unique special layer names
        }
        }
    
        generateAndSaveCSV(
        comps,
        uniqueTextLayerNames,
        uniqueFileLayerNames,
        uniqueSpecialLayerNames
        );
    }
    
    function generateAndSaveCSV(
        comps,
        uniqueTextLayerNames,
        uniqueFileLayerNames,
        uniqueSpecialLayerNames
    ) {
        var headers = ["Composition Name"];
    
        // Add headers for text layers
        for (var layerName in uniqueTextLayerNames) {
        if (uniqueTextLayerNames.hasOwnProperty(layerName)) {
            headers.push(layerName);
        }
        }
    
        // Add headers for file layers
        for (layerName in uniqueFileLayerNames) {
        if (uniqueFileLayerNames.hasOwnProperty(layerName)) {
            headers.push(layerName);
        }
        }
    
        // Now include headers for special layers starting with #
        for (layerName in uniqueSpecialLayerNames) {
        if (uniqueSpecialLayerNames.hasOwnProperty(layerName)) {
            headers.push(layerName); // Include headers for special layers
        }
        }
    
        var csvContent = headers.join(",") + "\n";
    
        for (var i = 0; i < comps.length; i++) {
        var compObj = comps[i];
        var compRow = ['"' + compObj.name.replace(/"/g, '""') + '"'];
    
        // Process each header to find corresponding content in the composition object
        for (var j = 1; j < headers.length; j++) {
            var header = headers[j];
            var isFilePath = header.indexOf(" File Path") > -1;
            var actualLayerName = isFilePath
            ? header.replace(" File Path", "")
            : header;
            var content = "";
    
            if (compObj.textLayersContent.hasOwnProperty(actualLayerName)) {
            content = compObj.textLayersContent[actualLayerName];
            } else if (compObj.fileLayersContent.hasOwnProperty(actualLayerName)) {
            content = compObj.fileLayersContent[actualLayerName];
            } else if (compObj.specialLayersContent.hasOwnProperty(actualLayerName)) {
            content = compObj.specialLayersContent[actualLayerName];
            }
    
            // Ensure content is handled correctly for CSV
            compRow.push('"' + (content ? content.replace(/"/g, '""') : "") + '"');
        }
    
        csvContent += compRow.join(",") + "\n";
        }
    
        saveCSVFile(csvContent);
    }
    
    function saveCSVFile(csvContent) {
        var file = new File(File.saveDialog("Save your CSV file", "*.csv"));
        if (file) {
        file.encoding = "UTF-8";
        file.open("w");
        if (file.write(csvContent)) {
            alert("CSV file saved successfully!");
        } else {
            alert("Failed to write to file.");
        }
        file.close();
        } else {
        alert("File save cancelled or failed to open file.");
        }
    }
    
    createUI();
}

function contentImport() {
    
        function chooseFilePath() {
            var file = File.openDialog("Select a TSV file", "TSV files:*.tsv;*.txt", false);
            if (file) {
                return file.fsName; // Return the full path of the file
            } else {
                alert("No file was selected.");
                return null; // Return null if no file is selected
            }
        }
        
        // Main function or appropriate place in the script to set the file path
        var filePath = chooseFilePath(); // Get the file path using the dialog box
        if (!filePath) {
            throw new Error("Script terminated: No file selected.");
        }
    
        
        // Function to read and parse the TSV file
        function parseDocument(filePath) {
            var file = new File(filePath);
            var data = [];
            if (file.open("r")) {
                var headers = file.readln().split("\t"); // Read header line
                alert("Headers: " + headers.join(", "));
                while (!file.eof) {
                    var line = file.readln();
                    var columns = line.split("\t");
                    var row = {};
                    for (var i = 0; i < headers.length; i++) {
                        row[headers[i]] = columns[i];
                    }
                    data.push(row);
                }
                file.close();
            } else {
                alert("Failed to open file: " + filePath);
            }
            return data;
        }
    
        // Function to import a file and place it in the specified folder
        function importFile(filePath, importFolder) {
            var file = new File(filePath);
            if (file.exists) {
                alert("File exists: " + filePath);
                var importOptions = new ImportOptions(file);
                if (importOptions.canImportAs(ImportAsType.FOOTAGE)) {
                    importOptions.importAs = ImportAsType.FOOTAGE;
                }
                var importedFile = app.project.importFile(importOptions);
                importedFile.parentFolder = importFolder; // Move the imported file to the specified folder
                return importedFile;
            } else {
                alert("File not found: " + filePath);
                return null;
            }
        }
    
        // Function to find a project item by name
        function findProjectItemByName(name) {
            for (var i = 1; i <= app.project.numItems; i++) {
                if (app.project.item(i).name === name) {
                    return app.project.item(i);
                }
            }
            return null;
        }
    
    
        // Function to update layers in a composition recursively
        function updateLayers(comp, rowData, precompsFolder, importFolder, importedFiles) {
            if (!importedFiles) {
                alert("importedFiles is undefined!");
                return;
            }
    
            for (var i = 1; i <= comp.layers.length; i++) {
                var layer = comp.layer(i);
                var layerName = layer.name;
    
                // Update text layers if the layer name matches any of the headers
                if (rowData[layerName]) {
                    var textValue = rowData[layerName];
                    if (textValue && layer.property("Source Text") != null) {
                        layer.property("Source Text").setValue(textValue);
                    }
                }
    
                // Replace layers with corresponding imported files if the layer name starts with ">"
                if (layerName.indexOf(">") === 0) {
                    var columnName = layerName.substring(1); // Remove the ">" symbol
                    var importedFile = importedFiles[columnName];
    
                    if (importedFile) {
                        layer.replaceSource(importedFile, false);
                        layer.enabled = true; // Ensure the layer is visible
                    } else {
                        alert("No imported file found for column: " + columnName);
                    }
                }
    
                // Replace layers with corresponding project items if the layer name starts with "#"
                if (layerName.indexOf("#") === 0) {
                    var columnName = layerName.substring(1); // Remove the "#" symbol
                    var projectItemName = rowData["#" + columnName]; // Get project item name from the TSV column
                    var projectItem = findProjectItemByName(projectItemName);
                    if (projectItem) {
                        layer.replaceSource(projectItem, false);
                        layer.enabled = true; // Ensure the layer is visible
                    } else {
                        alert("No project item found for name: " + projectItemName);
                    }
                }
    
                // Recursively update and duplicate text layers in precompositions
                if (layer.source instanceof CompItem) {
                    var preCompDuplicate = layer.source.duplicate();
                    preCompDuplicate.name = comp.name + " - " + layerName;
                    preCompDuplicate.parentFolder = precompsFolder;
    
                    // Update layers in the duplicated precomp
                    updateLayers(preCompDuplicate, rowData, precompsFolder, importFolder, importedFiles);
                    
    
                    // Update the layer to use the new precomp
                    layer.replaceSource(preCompDuplicate, false);
                    layer.enabled = true; // Ensure the layer is visible
                }
            }
        }
    
        // Function to create a comp from data
        function createCompFromData(rowData, mainComp, outputFolder, precompsFolder, importFolder) {
            var newComp = mainComp.duplicate();
            newComp.name = rowData["Composition Name"];
    
            var importedFiles = {};
    
            // Import files and store them in the dictionary
            for (var key in rowData) {
                if (key.indexOf(">") === 0) {
                    var filePath = rowData[key];
                    if (filePath) {
                        var importedFile = importFile(filePath, importFolder);
                        if (importedFile) {
                            importedFiles[key.substring(1)] = importedFile; // Store the imported file with the column name as key
                        }
                    }
                }
            }
    
            // Debugging log
            $.writeln("Imported Files: " + JSON.stringify(importedFiles));
    
            // Update layers in the duplicated composition and its precompositions
            updateLayers(newComp, rowData, precompsFolder, importFolder, importedFiles);
    
            // Move the new composition to the specified folder
            newComp.parentFolder = outputFolder;    
            return newComp;
        }
    
        // Main function
        function main() {
            var outputFolderName = "Generated Comps"; // Name of the folder for new compositions
            var precompsFolderName = "Precomps"; // Name of the folder for new precompositions
            var importFolderName = "Imported Files"; // Name of the folder for imported files
    
            var data = parseDocument(filePath);
            var messages = [];
    
            // Ensure the project has items
            if (app.project.items.length === 0) {
                var message = "No items in the project. Please ensure your project has at least one composition.";
                alert(message);
                throw new Error(message);
            }
    
            // Create a new folder for the generated compositions
            var outputFolder = app.project.items.addFolder(outputFolderName);
    
            // Create a nested folder for the precompositions within the output folder
            var precompsFolder = outputFolder.items.addFolder(precompsFolderName);
    
            // Create a folder for imported files within the output folder
            var importFolder = outputFolder.items.addFolder(importFolderName);
    
            for (var i = 0; i < data.length; i++) {
                var compName = data[i]["Composition Name"];
                if (compName) {
                    var mainComp = null;
                    for (var j = 1; j <= app.project.items.length; j++) {
                        if (app.project.items[j] instanceof CompItem && app.project.items[j].name === compName) {
                            mainComp = app.project.items[j];
                            break;
                        }
                    }
    
                    if (mainComp) {
                        createCompFromData(data[i], mainComp, outputFolder, precompsFolder, importFolder);
                    } else {
                        var message = "Composition not found: " + compName;
                        alert(message);
                        messages.push(message);
                    }
                }
            }
    
            // Show all messages at the end
            var successMessage = "Script completed successfully.\n\n" + messages.join("\n");
            alert(successMessage);
        }
    
        main();
     
}
