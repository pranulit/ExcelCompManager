
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

// Define a function to retrieve path data from a layer
function getPathData(layer) {
    if (layer.name.indexOf(">") === 0) {
        try {
            if (layer instanceof ShapeLayer && layer.property("ADBE Vector Shape - Group")) {
                var shapeProperty = layer.property("ADBE Vector Shape - Group");
                return shapeProperty.value.toString(); // Return the path data directly
            } else {
                // Return a default message when no path data is found
                return "no file path found";
            }
        } catch (error) {
            alert("Error in getPathData for " + layer.name + ": " + error.toString());
            return "no file path found"; // Return default message on error
        }
    } else {
        return "no file path found"; // Return default message if layer name does not start with '>'
    }
}




// Modify the searchPrecomps function to include alert statements for debugging
function searchPrecomps(comp, parentCompTextLayers, parentCompPathLayers) {
    for (var k = 1; k <= comp.numLayers; k++) {
        var layer = comp.layer(k);

        // If the layer is a composition itself, recurse into it
        if (layer.source instanceof CompItem) {
            searchPrecomps(layer.source, parentCompTextLayers, parentCompPathLayers);
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

        if (layer.name.indexOf(">") === 0) {
            var layerName = layer.name.substring(1); // Remove ">" symbol from the layer name
            var pathData = getPathData(layer); // Retrieve path data from the layer
            parentCompPathLayers[layerName] = pathData; // Store path data by layer name
            alert("Path data for layer " + layerName + ": " + pathData); // Informative alert on what was stored
        }
    }
}


function generateAndSaveCSV(comps, uniqueTextLayerNames, uniquePathLayerNames) {
    var headers = ["Composition Name"];
    var textLayerName, pathLayerName, layerName, textContent, pathContent;
    
    // Add text layer names to headers
    for (textLayerName in uniqueTextLayerNames) {
        if (uniqueTextLayerNames.hasOwnProperty(textLayerName)) {
            headers.push(textLayerName);
        }
    }

    // Add path layer names to headers
    for (pathLayerName in uniquePathLayerNames) {
        if (uniquePathLayerNames.hasOwnProperty(pathLayerName)) {
            headers.push(pathLayerName);
        }
    }

    var csvContent = headers.join(",") + "\n";

    // Iterate over comps with a traditional for loop
    for (var i = 0; i < comps.length; i++) {
        var compObj = comps[i];
        alert("Comp Object: " + JSON.stringify(compObj));
        var compRow = ['"' + compObj.name.replace(/"/g, '""') + '"']; // Start row with composition name

        // Fill row with data for each text layer or empty if none
        for (var j = 1; j < headers.length; j++) {
            layerName = headers[j];
            textContent = compObj.textLayersContent[layerName];
            pathContent = compObj.pathLayersContent[layerName];

            // Debugging alerts to check the values
            alert("Layer Name: " + layerName);
            alert("Text Content: " + textContent);
            alert("Path Content: " + pathContent);

            // Check if textContent is undefined or null
            if (textContent == null) {
                alert("Text Content is null or undefined");
                textContent = '';
            }

            // Check if pathContent is undefined or null
            if (pathContent == null) {
                alert("Path Content is null or undefined");
                pathContent = '';
            }

            // Debugging alert to check the final content being pushed to the row
            alert("Final Content: " + (textContent || pathContent).replace(/"/g, '""'));

            // Push the content to the row
            compRow.push('"' + (textContent || pathContent).replace(/"/g, '""') + '"');
        }

        // Debugging alerts to check compRow and headers length
        alert("Comp Row: " + JSON.stringify(compRow));
        alert("Headers Length: " + headers.length);

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
