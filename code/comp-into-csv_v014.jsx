function exportCompositionNamesToCSV() {
    var compsInfo = {};
    var uniqueTextLayerNames = {};
    var uniqueLayerPaths = {};

    // Collect all compositions and prepare the data structure
    for (var i = 1; i <= app.project.numItems; i++) {
        var item = app.project.item(i);
        if (item instanceof CompItem) {
            compsInfo[item.name] = {
                name: item.name,
                textLayersContent: [],
                layerPaths: [],
                usedAsNested: false
            };
        }
    }

    // Process each composition to collect text and path data
    for (var compName in compsInfo) {
        if (compsInfo.hasOwnProperty(compName)) {
            var comp = getCompByName(compName);
            if (comp) {
                processCompLayers(comp, compName, compsInfo, uniqueTextLayerNames, uniqueLayerPaths);
            }
        }
    }

    // Prepare headers for the CSV
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

    // Sort the compositions based on the number of text layers
    var textLayerCounts = {};
    for (var compName in compsInfo) {
        if (compsInfo.hasOwnProperty(compName)) {
            var compObj = compsInfo[compName];
            textLayerCounts[compName] = compObj.textLayersContent.length;
        }
    }

    // Populate sortedCompNames array
    var sortedCompNames = [];
    for (var compName in textLayerCounts) {
        if (textLayerCounts.hasOwnProperty(compName)) {
            sortedCompNames.push(compName);
        }
    }

    // Sort sortedCompNames array based on the number of text layers
    sortedCompNames.sort(function(a, b) {
        return textLayerCounts[b] - textLayerCounts[a];
    });

    // Initialize CSV content with headers
    var csvContent = headers.join(",") + "\n";

    // Build the CSV content for each composition
    for (var compIndex = 0; compIndex < sortedCompNames.length; compIndex++) {
        var compName = sortedCompNames[compIndex];
        var compObj = compsInfo[compName];
        var compRow = new Array(headers.length);
        for (var i = 0; i < headers.length; i++) {
            compRow[i] = '""'; // Manually fill the array with empty values
        }

        compRow[0] = "\"" + compObj.name.replace(/"/g, '""') + "\""; // Set composition name

        // Fill text layer content
        for (var j = 0; j < compObj.textLayersContent.length; j++) {
            var textLayer = compObj.textLayersContent[j];
            var textLayerName = textLayer.name.split("^").slice(1).join("^");
            // Find the index of textLayerName in headers array
            var headerIndex = -1;
            for (var i = 0; i < headers.length; i++) {
                if (headers[i] === textLayerName) {
                    headerIndex = i;
                    break;
                }
            }
            if (headerIndex !== -1) {
                compRow[headerIndex] = "\"" + textLayer.text.replace(/"/g, '""') + "\"";
            }
        }

        // Fill layer paths
        for (var j = 0; j < compObj.layerPaths.length; j++) {
            var layerPath = compObj.layerPaths[j];
            var headerIndex = -1;
            for (var k = 0; k < headers.length; k++) {
                if (headers[k] === "Layer Path: " + layerPath) {
                    headerIndex = k;
                    break;
                }
            }
            if (headerIndex !== -1) {
                compRow[headerIndex] = "\"" + layerPath.replace(/"/g, '""') + "\"";
            }
        }


        // Join the row with commas and add to CSV content
        csvContent += compRow.join(",") + "\n";
    }

    // Save the CSV content to a file
    var saveCSVFile = function(csvContent) {
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
    };

    saveCSVFile(csvContent);
}

// Helper function to get a composition by name
function getCompByName(compName) {
    for (var i = 1; i <= app.project.numItems; i++) {
        var item = app.project.item(i);
        if (item instanceof CompItem && item.name === compName) {
            return item;
        }
    }
    return null; // Return null if no composition matches the name
}

exportCompositionNamesToCSV();