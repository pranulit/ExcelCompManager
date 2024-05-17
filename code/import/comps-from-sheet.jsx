{
  // Function to read TSV file
  function readTSV(filePath) {
    var file = new File(filePath);
    if (!file.exists) {
      $.writeln("TSV file not found: " + filePath);
      throw new Error("TSV file not found: " + filePath);
    }

    file.open("r");
    var data = [];
    var headers;
    while (!file.eof) {
      var line = file.readln();
      var values = line.split("\t");

      if (!headers) {
        headers = values; // first row as header
        $.writeln("Headers: " + headers.join(", "));
      } else {
        var row = {};
        for (var i = 0; i < headers.length; i++) {
          row[headers[i]] = values[i];
        }
        data.push(row);
      }
    }
    file.close();
    $.writeln("TSV file read successfully: " + filePath);
    return data;
  }

  // Function to import a file and place it in the specified folder
  function importFile(filePath, importFolder) {
    var file = new File(filePath);
    $.writeln("Checking file: " + filePath);
    if (file.exists) {
      $.writeln("File exists: " + filePath);
      var importOptions = new ImportOptions(file);
      if (importOptions.canImportAs(ImportAsType.FOOTAGE)) {
        importOptions.importAs = ImportAsType.FOOTAGE;
      }
      var importedFile = app.project.importFile(importOptions);
      $.writeln("File imported: " + filePath);
      importedFile.parentFolder = importFolder; // Move the imported file to the specified folder
      $.writeln("File moved to folder: " + importFolder.name);
      return importedFile;
    } else {
      $.writeln("File not found: " + filePath);
      return null; // Handle file not found by returning null
    }
  }

  // Function to update layers in a composition recursively
  function updateLayers(comp, rowData, precompsFolder, importFolder) {
    for (var i = 1; i <= comp.layers.length; i++) {
      var layer = comp.layer(i);
      var layerName = layer.name;
      $.writeln("Processing layer: " + layerName);

      // Update text layers if the layer name matches any of the headers
      if (rowData[layerName]) {
        var textValue = rowData[layerName];
        if (textValue && layer.property("Source Text") != null) {
          layer.property("Source Text").setValue(textValue);
          $.writeln(
            "Updated text for layer: " + layerName + " to: " + textValue
          );
        }
      }

      // Import and assign files to layers if the layer name starts with ">"
      if (layerName.indexOf(">") === 0) {
        var columnName = layerName.substring(1); // Remove the ">" symbol
        var filePath = rowData[columnName];
        if (filePath) {
          $.writeln(
            "Found file path for layer: " + layerName + " - " + filePath
          );
          var importedFile = importFile(filePath, importFolder);
          if (importedFile) {
            layer.replaceSource(importedFile, false);
            layer.enabled = true; // Ensure the layer is visible
            $.writeln(
              "Replaced source for layer: " +
                layer.name +
                " with imported file: " +
                filePath
            );
          } else {
            $.writeln("Failed to import file for layer: " + layerName);
          }
        }
      }

      // Recursively update text layers in precompositions
      if (layer.source instanceof CompItem) {
        $.writeln("Found precomposition layer: " + layerName);
        updateLayers(layer.source, rowData, precompsFolder, importFolder);
      }
    }
  }

  // Function to create a comp from data
  function createCompFromData(
    rowData,
    compIndex,
    mainComp,
    outputFolder,
    precompsFolder,
    importFolder
  ) {
    var newComp = mainComp.duplicate();
    newComp.name = rowData["Composition Name"] || "Comp_" + compIndex;
    $.writeln("Created new comp: " + newComp.name);

    // Update layers in the duplicated composition and its precompositions
    updateLayers(newComp, rowData, precompsFolder, importFolder);

    // Move the new composition to the specified folder
    newComp.parentFolder = outputFolder;
    $.writeln("Moved new comp to folder: " + outputFolder.name);

    return newComp;
  }

  // Main function
  function main() {
    var tsvPath =
      "/Users/tadaspranulis/Desktop/Code/after-effects/AE-content-extractor/excel/02_txt/text-path-name_extract_v001.txt"; // Path to the TSV file
    var outputFolderName = "Generated Comps"; // Name of the folder for new compositions
    var precompsFolderName = "Precomps"; // Name of the folder for new precompositions
    var importFolderName = "Imported Files"; // Name of the folder for imported files

    var data = readTSV(tsvPath);

    // Ensure the project has items
    if (app.project.items.length === 0) {
      $.writeln(
        "No items in the project. Please ensure your project has at least one composition."
      );
      throw new Error(
        "No items in the project. Please ensure your project has at least one composition."
      );
    }

    // Find the main composition (assuming it is the first composition in the project)
    var mainComp = null;
    for (var i = 1; i <= app.project.items.length; i++) {
      if (app.project.items[i] instanceof CompItem) {
        mainComp = app.project.items[i];
        break;
      }
    }

    if (!mainComp) {
      $.writeln(
        "No composition found in the project. Please ensure your project has at least one composition."
      );
      throw new Error(
        "No composition found in the project. Please ensure your project has at least one composition."
      );
    }

    $.writeln("Main composition found: " + mainComp.name);

    // Create a new folder for the generated compositions
    var outputFolder = app.project.items.addFolder(outputFolderName);
    $.writeln("Created folder: " + outputFolderName);

    // Create a nested folder for the precompositions within the output folder
    var precompsFolder = outputFolder.items.addFolder(precompsFolderName);
    $.writeln("Created folder: " + precompsFolderName);

    // Create a folder for imported files within the output folder
    var importFolder = outputFolder.items.addFolder(importFolderName);
    $.writeln("Created folder: " + importFolderName);

    for (var i = 0; i < data.length; i++) {
      $.writeln("Processing row " + (i + 1));
      createCompFromData(
        data[i],
        i + 1,
        mainComp,
        outputFolder,
        precompsFolder,
        importFolder
      );
    }

    $.writeln("Script completed successfully.");
  }

  main();
}
