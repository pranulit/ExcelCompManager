
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
    var dlg = new Window("dialog", "Select Compositions", undefined, {resizeable: false});
    dlg.size = [400, 400];

    // Add a group for the search bar
    var searchGroup = dlg.add("group");
    searchGroup.orientation = "row";
    // Add a text edit field and a search button to the group
    var searchEdit = searchGroup.add("edittext", undefined, "");
    searchEdit.characters = 20;
    var searchButton = searchGroup.add("button", undefined, "Search");

    // Add a list box to display items with multi-selection enabled
    var listBox = dlg.add("listbox", [0, 0, 380, 300], [], {multiselect: true});
    listBox.alignment = "fill";

    // Populate the list box initially
    populateListBox(listBox, "");  // Assuming a function populateListBox is defined

    // Define the search button click behavior
    searchButton.onClick = function() {
        populateListBox(listBox, searchEdit.text);
    };

    // Add a group for buttons
    var buttons = dlg.add("group");
    buttons.alignment = "right";
    var cancelButton = buttons.add("button", undefined, "Cancel", { name: "cancel" });
    var saveButton = buttons.add("button", undefined, "Save", { name: "ok" });

    // Define the cancel button behavior
    cancelButton.onClick = function() { dlg.close(); };
    // Define the save button behavior
    saveButton.onClick = function() {
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


function searchPrecomps(comp, parentCompTextLayers) {
    for (var k = 1; k <= comp.numLayers; k++) {
        var layer = comp.layer(k);

        // If the layer is a composition itself, recurse into it
        if (layer.source instanceof CompItem) {
            searchPrecomps(layer.source, parentCompTextLayers);
        }

        // Check if the layer is a text layer and has the "Source Text" property
        if (layer instanceof TextLayer && layer.property("Source Text") != null) {
            var textSource = layer.property("Source Text").value;
            if (textSource) {
                var textContent = textSource.text.replace(/[\r\n]+/g, " "); // Clean up new lines
                var layerName = layer.name.split("^").slice(1).join("^"); // Assuming your layer names might contain special markers or prefixes
                parentCompTextLayers[layerName] = textContent; // Store text content by layer name
            }
        }
    }
}

function exportSelectedCompositions(compCheckboxes) {
    var comps = [];
    var uniqueTextLayerNames = {};

    for (var i = 0; i < compCheckboxes.length; i++) {
        var compItem = compCheckboxes[i];
        if (!compItem || !compItem.name || !compItem.numLayers || !compItem.layer) {
            continue; // Skip if compItem is invalid
        }

        // Reset text layer data storage for this composition
        var parentCompTextLayers = {};

        // Call searchPrecomps to extract text layers from parent comp, pass the parentCompTextLayers object to store results
        searchPrecomps(compItem, parentCompTextLayers);

        // Object to represent the composition with its text layers
        var compObj = {
            name: compItem.name,
            textLayersContent: parentCompTextLayers // Assign extracted text layers
        };

        comps.push(compObj);
        
        // Update unique text layer names dictionary
        for (var layerName in parentCompTextLayers) {
            uniqueTextLayerNames[layerName] = true;
        }
    }

    generateAndSaveCSV(comps, uniqueTextLayerNames);
}

function generateAndSaveCSV(comps, uniqueTextLayerNames) {
    var headers = ["Composition Name"];
    // Manually iterate through the uniqueTextLayerNames to add them to headers
    for (var layerName in uniqueTextLayerNames) {
        if (uniqueTextLayerNames.hasOwnProperty(layerName)) {
            headers.push(layerName);
        }
    }

    var csvContent = headers.join(",") + "\n";

    // Iterate over comps with a traditional for loop
    for (var i = 0; i < comps.length; i++) {
        var compObj = comps[i];
        var compRow = ['"' + compObj.name.replace(/"/g, '""') + '"']; // Start row with composition name

        // Iterate over headers to fill row with data for each text layer or empty if none
        for (var j = 1; j < headers.length; j++) {
            var layerName = headers[j];
            var textContent = compObj.textLayersContent[layerName];
            compRow.push(textContent ? '"' + textContent.replace(/"/g, '""') + '"' : '""');
        }

        // Join row data and add to CSV content
        csvContent += compRow.join(",") + "\n";
    }

    // Save CSV file with the generated content
    saveCSVFile(csvContent);
}

function saveCSVFile(csvContent) {
    var file = new File(File.saveDialog("Save your CSV file", "*.csv"));
    if (file) {
        file.encoding = "UTF-8";
        file.open("w");
        var writeSuccess = file.write(csvContent);
        file.close();
        if (writeSuccess) {
            alert("CSV file saved successfully!");
        } else {
            alert("Failed to write to file.");
        }
    } else {
        alert("File save cancelled or failed to open file.");
    }
}

createUI();
