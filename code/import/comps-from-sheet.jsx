{
    var filePath = "C:\\Users\\TAPE-B\\Desktop\\Code\\AE-content-extraction\\excel\\txt\\00_main.txt"; // Path to the TSV file

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
        alert("Checking file: " + filePath);
        if (file.exists) {
            alert("File exists: " + filePath);
            var importOptions = new ImportOptions(file);
            if (importOptions.canImportAs(ImportAsType.FOOTAGE)) {
                importOptions.importAs = ImportAsType.FOOTAGE;
            }
            var importedFile = app.project.importFile(importOptions);
            importedFile.parentFolder = importFolder; // Move the imported file to the specified folder
            alert("File imported and moved to folder: " + importFolder.name);
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
            alert("Processing layer: " + layerName);

            // Update text layers if the layer name matches any of the headers
            if (rowData[layerName]) {
                var textValue = rowData[layerName];
                if (textValue && layer.property("Source Text") != null) {
                    layer.property("Source Text").setValue(textValue);
                    alert("Updated text for layer: " + layerName + " to: " + textValue);
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
                alert("Found precomposition layer: " + layerName);
                var preCompDuplicate = layer.source.duplicate();
                preCompDuplicate.name = comp.name + " - " + layerName;
                preCompDuplicate.parentFolder = precompsFolder;

                // Update layers in the duplicated precomp
                updateLayers(preCompDuplicate, rowData, precompsFolder, importFolder, importedFiles);
                

                // Update the layer to use the new precomp
                layer.replaceSource(preCompDuplicate, false);
                layer.enabled = true; // Ensure the layer is visible
                alert("Duplicated precomp: " + preCompDuplicate.name);
            }
        }
    }

    // Function to create a comp from data
    function createCompFromData(rowData, mainComp, outputFolder, precompsFolder, importFolder) {
        var newComp = mainComp.duplicate();
        newComp.name = rowData["Composition Name"];
        alert("Created new comp: " + newComp.name);

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
        alert("Moved new comp to folder: " + outputFolder.name);

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
        alert("Created folder: " + outputFolderName);

        // Create a nested folder for the precompositions within the output folder
        var precompsFolder = outputFolder.items.addFolder(precompsFolderName);
        alert("Created folder: " + precompsFolderName);

        // Create a folder for imported files within the output folder
        var importFolder = outputFolder.items.addFolder(importFolderName);
        alert("Created folder: " + importFolderName);

        for (var i = 0; i < data.length; i++) {
            alert("Processing row " + (i + 1));
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
