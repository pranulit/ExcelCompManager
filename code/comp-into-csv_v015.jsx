function indexOf(array, searchElement) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] === searchElement) {
            return i;
        }
    }
    return -1;
}

function processCompLayers(comp, compName, compsInfo, uniqueTextLayerNames, uniqueLayerPaths) {
    for (var i = 1; i <= comp.numLayers; i++) {
        var layer = comp.layer(i);
        if (layer instanceof TextLayer) {
            var textProp = layer.property("Source Text");
            var textValue = textProp.value;
            compsInfo[compName].textLayersContent.push({
                name: layer.name,
                text: textValue.text
            });
            uniqueTextLayerNames[layer.name] = true; // Record unique text layer names
        }
        if (layer instanceof ShapeLayer) {
            var contents = layer.property("Contents");
            for (var j = 1; j <= contents.numProperties; j++) {
                var prop = contents.property(j);
                if (prop.propertyType === PropertyType.INDEXED_GROUP && prop.matchName === 'ADBE Vector Shape - Group') {
                    var pathProp = prop.property("ADBE Vector Shape");
                    compsInfo[compName].layerPaths.push(pathProp.name);
                    uniqueLayerPaths[pathProp.name] = true; // Record unique path names
                }
            }
        }
    }
}

function exportCompositionNamesToCSV() {
    var compsInfo = {};
    var uniqueTextLayerNames = {};
    var uniqueLayerPaths = {};

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

    for (var compName in compsInfo) {
        if (compsInfo.hasOwnProperty(compName)) {
            var comp = getCompByName(compName);
            if (comp) {
                processCompLayers(comp, compName, compsInfo, uniqueTextLayerNames, uniqueLayerPaths);
            }
        }
    }

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
    var compNames = [];
    for (var compName in compsInfo) {
        if (compsInfo.hasOwnProperty(compName)) {
            compNames.push(compName);
        }
    }

    compNames.sort(function(a, b) {
        return compsInfo[a].textLayersContent.length - compsInfo[b].textLayersContent.length;
    });

    for (var compIndex = 0; compIndex < compNames.length; compIndex++) {
        var compName = compNames[compIndex];
        var compObj = compsInfo[compName];
        var compRow = new Array(headers.length);
        for (var i = 0; i < headers.length; i++) {
            compRow[i] = '""';
        }
        compRow[0] = "\"" + compObj.name.replace(/"/g, '""') + "\"";

        for (var j = 0; j < compObj.textLayersContent.length; j++) {
            var textLayer = compObj.textLayersContent[j];
            var textLayerName = textLayer.name.split("^").slice(1).join("^");
            var headerIndex = indexOf(headers, textLayerName);
            if (headerIndex !== -1) {
                compRow[headerIndex] = "\"" + textLayer.text.replace(/"/g, '""') + "\"";
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

function getCompByName(compName) {
    for (var i = 1; i <= app.project.numItems; i++) {
        var item = app.project.item(i);
        if (item instanceof CompItem && item.name === compName) {
            return item;
        }
    }
    return null;
}

exportCompositionNamesToCSV();
