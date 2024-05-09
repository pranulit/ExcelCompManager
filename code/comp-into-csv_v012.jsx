function indexOf(array, element) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] === element) {
            return i;
        }
    }
    return -1; // Element not found
}


function exportCompositionNamesToCSV() {
    var comps = []; // Array to store composition objects
    var uniqueTextLayerNames = {};
    var uniqueLayerPaths = {}; // Store unique paths for layers with '>'

    for (var i = 1; i <= app.project.numItems; i++) {
        var item = app.project.item(i);
        if (item instanceof CompItem) { // Check if the item is a CompItem
            var compObj = {
                name: item.name, // Composition's name
                textLayersContent: [], // Array to store content from text layers
                layerPaths: [] // Array to store paths of layers with '>'
            };

            for (var j = 1; j <= item.numLayers; j++) {
                var layer = item.layer(j);
                if (layer instanceof TextLayer && layer.property("Source Text") != null) {
                    var textSource = layer.property("Source Text").value;
                    if (textSource) {
                        var textContent = textSource.text.replace(/[\r\n]+/g, " ");
                        compObj.textLayersContent.push({text: textContent, name: layer.name});
                        var layerName = layer.name.split("^").slice(1).join("^");
                        uniqueTextLayerNames[layerName] = true;
                    }
                }
                if (layer.name.indexOf(">") !== -1) {
                    var path = item.name + "/" + layer.name;
                    compObj.layerPaths.push(path);
                    uniqueLayerPaths[path] = true;
                }
            }
            comps.push(compObj);
        }
    }

    // CSV Generation

    //HEADERS
    var headers = ["Composition Name"];
    for (var layerName in uniqueTextLayerNames) {
        if (uniqueTextLayerNames.hasOwnProperty(layerName)) {
            headers.push(layerName);
        }
    }
    for (var path in uniqueLayerPaths) {
        if (uniqueLayerPaths.hasOwnProperty(path)) {
            headers.push("Layer Path: " + path);
        }
    }

    var csvContent = headers.join(",") + "\n";

    for (var k = 0; k < comps.length; k++) {
        var compObj = comps[k];
        var compRow = new Array(headers.length);
        compRow[0] = "\"" + compObj.name.replace(/"/g, '""') + "\"";

        for (var h = 1; h < headers.length; h++) {
            compRow[h] = '""';
        }

        for (var j = 0; j < compObj.textLayersContent.length; j++) {
            var textLayer = compObj.textLayersContent[j];
            var textLayerName = textLayer.name.split("^").slice(1).join("^");
            var headerIndex = indexOf(headers, textLayerName);

            if (headerIndex !== -1) {
                var textContent = textLayer.text.replace(/[\r\n]+/g, " ").replace(/"/g, '""');
                compRow[headerIndex] = "\"" + textContent + "\"";
            }
        }

        for (var j = 0; j < compObj.layerPaths.length; j++) {
            var layerPath = compObj.layerPaths[j];
            var headerIndex = indexOf(headers, "Layer Path: " + layerPath);
            if (headerIndex !== -1) {
                compRow[headerIndex] = "\"" + layerPath.replace(/"/g, '""') + "\"";
            }
        }

        csvContent += compRow.join(",") + "\n";
    }

    // Function to save the CSV content to a file
    var saveCSVFile = function(csvContent) {
        var file = new File(File.saveDialog("Save your CSV file", "*.csv"));
        if (file) {
            file.encoding = "UTF-8"; // Set encoding
            file.open("w"); // Open file for writing
            var writeSuccess = file.write(csvContent); // Attempt to write and store result
            file.close(); // Close the file

            if (writeSuccess) {
                alert("CSV file saved successfully!");
            } else {
                alert("Failed to write to file.");
            }
        } else {
            alert("File save cancelled or failed to open file.");
        }
    };

    saveCSVFile(csvContent); // Call the function to save the CSV
}

exportCompositionNamesToCSV();