(function() {
    var project = app.project;
    if (!project) {
        alert("No project found.");
        return;
    }

    function indexOfArray(array, searchElement) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === searchElement) {
                return i;
            }
        }
        return -1; // Return -1 if the element is not found
    }

    var compData = [];
    var usedAsPrecomp = {};
    // Initialize an empty array for storing unique text layer names
    var textNameArray = [];
    var hashLayers = []; // Initialize an array for BG layer names
    var filePathArray = [];

    app.beginUndoGroup("Export Comp Text Layers to CSV");

    function searchPrecomps(comp, layers) {
        for (var k = 1; k <= comp.numLayers; k++) {
            var layer = comp.layer(k);
    
            if (layer.name.indexOf("#") !== -1) {
                var bgLayerName = layer.name; // Keep original naming
                var cleanBgLayerName = bgLayerName.replace("#", ""); // Remove the first occurrence of "#"
    
                // Check if we already have this bgLayerName in our list; if not, add it
                if (indexOfArray(hashLayers, cleanBgLayerName) === -1) {
                    hashLayers.push(cleanBgLayerName);
                }
    
                // Map the cleanBgLayerName to the source layer's name within the current composition
                layers[cleanBgLayerName] = layer.source ? layer.source.name : "Source not found";
            }

                        
            // Adjusted handling for ^ layers to capture actual text content
            if (layer.name.indexOf("^") !== -1 && layer.property("sourceText") !== undefined) {
                var textName = layer.name; // Keep original naming
                var cleanTextName = textName.replace("^", ""); // Remove the first occurrence of ^

                // Check if we already have this textName in our list; if not, add it
                if (indexOfArray(textNameArray, cleanTextName) === -1) {
                    textNameArray.push(cleanTextName);
                }

                // Retrieve and map the actual text content of the text layer
                var textProperty = layer.property("sourceText");
                if (textProperty != null) {
                    var textDocument = textProperty.value;  // TextDocument object
                    if (textDocument.text !== undefined) {
                        layers[cleanTextName] = textDocument.text;  // Retrieve the text string
                    } else {
                        layers[cleanTextName] = "Text content not available";
                    }
                } else {
                    layers[cleanTextName] = "Not a valid text layer";
                }
            }




            // New handling for layers starting with >
            if (layer.name.indexOf(">") !== -1) {
                var filePathLayerName = layer.name; // Keep original naming
                var cleanFilePathLayerName = filePathLayerName.replace(">", ""); // Remove the first occurrence of ">"
                
                // Check if we already have this cleanFilePathLayerName in our list; if not, add it
                // Use the custom indexOfArray function for ECMAScript 3 compatibility
                if (indexOfArray(filePathArray, cleanFilePathLayerName) === -1) {
                    filePathArray.push(cleanFilePathLayerName);
                }
                
                // Retrieve and map the source file path of the layer
                if (layer.source && layer.source.file) {
                    layers[cleanFilePathLayerName] = layer.source.file.fsName; // fsName gives the full path
                } else {
                    layers[cleanFilePathLayerName] = "File path not found";
                }
            }


    
            // Recurse into nested precomps
            if (layer.source instanceof CompItem && /^!/.test(layer.name)) {
                searchPrecomps(layer.source, layers);
            }
        }

    }

    // First Pass: Identify precomps
    for (var i = 1; i <= project.numItems; i++) {
        var item = project.item(i);
        if (item instanceof CompItem) {
            for (var j = 1; j <= item.numLayers; j++) {
                var layer = item.layer(j);
                if (layer.source instanceof CompItem) {
                    // Mark this composition as used as a precomp
                    usedAsPrecomp[layer.source.id] = true;
                }
            }
        }
    }

    // Second Pass: Generate compData excluding precomps that are used in other comps
    for (var i = 1; i <= project.numItems; i++) {
        var item = project.item(i);
        if (item instanceof CompItem && !usedAsPrecomp[item.id]) {
            var comp = item;
            var layers = {};

            searchPrecomps(comp, layers);
            compData.push({compName: comp.name, layers: layers});
        }
    }


textNameArray.sort(function(a, b) {
    // Split the names into text and numeric parts
    var partsA = a.split(/_(?=\d+$)/); // Splits at underscore before digits at the end
    var textPartA = partsA[0];
    var numPartA = parseInt(partsA[1], 10) || 0; // Use 0 if there's no numeric part
    
    var partsB = b.split(/_(?=\d+$)/); // Splits at underscore before digits at the end
    var textPartB = partsB[0];
    var numPartB = parseInt(partsB[1], 10) || 0; // Use 0 if there's no numeric part

    // Compare the text parts
    if (textPartA < textPartB) return -1;
    if (textPartA > textPartB) return 1;

    // If text parts are equal, compare the numeric parts
    return numPartA - numPartB;
});

function filterArray(array, test) {
    var result = [];
    for (var i = 0; i < array.length; i++) {
        if (test(array[i], i, array)) {
            result.push(array[i]);
        }
    }
    return result;
}

// Filter out compositions starting with "!" before generating the CSV
var filteredCompData = filterArray(compData, function(comp) {
    // Use indexOf to check if the compName does not start with "!"
    return comp.compName.indexOf("!") !== 0;
});


// CSV Generation
var csvContent = "COMP";

// Append sorted column names for each type
for (var i = 0; i < textNameArray.length; i++) {
    csvContent += "," + textNameArray[i];
}
for (var i = 0; i < filePathArray.length; i++) {
    csvContent += "," + filePathArray[i];
}
for (var i = 0; i < hashLayers.length; i++) {
    csvContent += "," + hashLayers[i];
}
csvContent += "\n";

// Now, loop over filteredCompData instead of compData to exclude comps starting with "!"
for (var i = 0; i < filteredCompData.length; i++) {
    var comp = filteredCompData[i];
    var rowData = "\"" + comp.compName + "\"";

    for (var j = 0; j < textNameArray.length; j++) {
        var name = textNameArray[j];
        rowData += "," + (comp.layers[name] ? "\"" + comp.layers[name] + "\"" : "");
    }

    for (var j = 0; j < filePathArray.length; j++) {
        var name = filePathArray[j];
        rowData += "," + (comp.layers[name] ? "\"" + comp.layers[name] + "\"" : "");
    }


    for (var j = 0; j < hashLayers.length; j++) {
        var name = hashLayers[j];
        rowData += "," + (comp.layers[name] ? "\"" + comp.layers[name] + "\"" : "");
    }

    csvContent += rowData + "\n";
}

app.endUndoGroup();

// Save Dialog
var file = File.saveDialog("Save CSV File", "*.csv");
if (file) {
    if (file.open("w")) {

        file.encoding = "UTF-8"; // Set encoding
        file.write(csvContent); // Corrected from 'data' to 'csvContent'
        file.close();
        alert("Data exported to " + file.fsName);
    } else {
        alert("Failed to open file for writing.");
    }
} else {
    alert("Export cancelled.");
}

})();