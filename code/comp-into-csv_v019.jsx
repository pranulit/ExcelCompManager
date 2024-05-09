function indexOf(array, element) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] === element) {
            return i;
        }
    }
    return -1; // Element not found
}



function createUI() {
    var dlg = new Window("dialog", "Select Compositions", undefined, {resizeable:false});
    dlg.size = [400, 400]; // Set dialog window size

    // Search input field and search button
    var searchGroup = dlg.add("group");
    searchGroup.alignment = "fill";
    var searchInput = searchGroup.add("edittext", undefined, "", {multiline:false});
    searchInput.characters = 20; // Set the maximum number of characters for the search input

    var searchButton = searchGroup.add("button", undefined, "Search");
    searchButton.onClick = filterList; // Trigger search when the button is clicked

    // Add a list box to contain the composition names
    var listBox = dlg.add("listbox", [0, 0, 400, 300], undefined, {multiselect:true});
    listBox.alignment = "fill";

    var allComps = []; // Array to hold all composition names
    var compItems = []; // Array to hold actual composition objects for export

    // Add composition names to the list box
    var project = app.project;
    var compositions = project.items;
    for (var i = 1; i <= compositions.length; i++) {
        var compItem = compositions[i];
        if (compItem instanceof CompItem) {
            listBox.add("item", compItem.name);
            allComps.push(compItem.name); // Store the name for filtering
            compItems.push(compItem); // Store the composition object for export
        }
    }

    function filterList() {
        var query = searchInput.text.toLowerCase();
        listBox.removeAll(); // Remove all items to repopulate

        for (var i = 0; i < allComps.length; i++) {
            if (allComps[i].toLowerCase().indexOf(query) !== -1) {
                listBox.add("item", allComps[i]); // Add only items that match the query
            }
        }
        
        dlg.layout.layout(true);
        dlg.layout.resize();
    }

    // Update the filter as you type
    searchInput.onChanging = filterList;

    // Buttons for actions
    var buttons = dlg.add("group");
    buttons.alignment = "right";
    var cancelButton = buttons.add("button", undefined, "Cancel", { name: "cancel" });
    var saveButton = buttons.add("button", undefined, "Save Text Content", { name: "ok" });

    cancelButton.onClick = function() { dlg.close(); };
    saveButton.onClick = function() {
        var selectedComps = [];
        for (var i = 0; i < listBox.items.length; i++) {
            if (listBox.items[i].selected && compItems[i]) { // Check that compItems[i] exists
                selectedComps.push(compItems[i]);
            }
        }
        try {
            exportSelectedCompositions(selectedComps);
        } catch (e) {
            alert("Error: " + e.toString());
        }
        dlg.close();
    };
    

    dlg.layout.layout(true);
    dlg.center();
    dlg.show();
}



function exportSelectedCompositions(compCheckboxes) {
    var comps = [];
    var uniqueTextLayerNames = {};

    for (var i = 0; i < compCheckboxes.length; i++) {
        var item = compCheckboxes[i].compItem;
        var compObj = {
            name: item.name,
            textLayersContent: []
        };

        for (var j = 1; j <= item.numLayers; j++) {
            var layer = item.layer(j);
            if (layer instanceof TextLayer && layer.property("Source Text") != null) {
                var textSource = layer.property("Source Text").value;
                if (textSource) {
                    var textContent = textSource.text.replace(/[\r\n]+/g, " ");
                    compObj.textLayersContent.push({ text: textContent, name: layer.name });
                    var layerName = layer.name.split("^").slice(1).join("^");
                    uniqueTextLayerNames[layerName] = true;
                }
            }
        }

        comps.push(compObj);
    }
    generateAndSaveCSV(comps, uniqueTextLayerNames);
}

function generateAndSaveCSV(comps, uniqueTextLayerNames) {
    var headers = ["Composition Name"];
    for (var layerName in uniqueTextLayerNames) {
        if (uniqueTextLayerNames.hasOwnProperty(layerName)) {
            headers.push(layerName);
        }
    }

    var csvContent = headers.join(",") + "\n";
    for (var k = 0; k < comps.length; k++) {
        var compObj = comps[k];
        var compRow = new Array(headers.length);
        for (var m = 0; m < compRow.length; m++) { compRow[m] = '""'; } // Initialize array elements
        compRow[0] = "\"" + compObj.name.replace(/"/g, '""') + "\"";

        for (var j = 0; j < compObj.textLayersContent.length; j++) {
            var textLayer = compObj.textLayersContent[j];
            var textLayerName = textLayer.name.split("^").slice(1).join("^");
            var headerIndex = indexOf(headers, textLayerName);
            if (headerIndex !== -1) {
                var textContent = textLayer.text.replace(/"/g, '""');
                compRow[headerIndex] = "\"" + textContent + "\"";
            }
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