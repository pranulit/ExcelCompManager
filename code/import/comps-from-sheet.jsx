{
  // Function to read TSV file
  function readTSV(filePath) {
    var file = new File(filePath);
    if (!file.exists) {
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
      } else {
        var row = {};
        for (var i = 0; i < headers.length; i++) {
          row[headers[i]] = values[i];
        }
        data.push(row);
      }
    }
    file.close();
    return data;
  }

  // Function to update text layers in a composition
  function updateTextLayers(comp, rowData, precompsFolder) {
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

      // Recursively update and duplicate text layers in precompositions
      if (layer.source instanceof CompItem) {
        var preComp = layer.source.duplicate();
        preComp.name = comp.name + " - " + layerName;
        preComp.parentFolder = precompsFolder;

        // Update text layers in the duplicated precomp
        updateTextLayers(preComp, rowData, precompsFolder);

        // Update the layer to use the new precomp
        layer.replaceSource(preComp, false);
      }
    }
  }

  // Function to create a comp from data
  function createCompFromData(
    rowData,
    compIndex,
    mainComp,
    outputFolder,
    precompsFolder
  ) {
    var newComp = mainComp.duplicate();
    newComp.name = rowData["Composition Name"] || "Comp_" + compIndex;

    // Update text layers in the duplicated composition and its precompositions
    updateTextLayers(newComp, rowData, precompsFolder);

    // Move the new composition to the specified folder
    newComp.parentFolder = outputFolder;

    return newComp;
  }

  // Main function
  function main() {
    var tsvPath =
      "/Users/tadaspranulis/Desktop/Code/after-effects/AE-content-extractor/excel/02_txt/text-path-name_extract_v001.txt"; // Path to the TSV file
    var outputFolderName = "Generated Comps"; // Name of the folder for new compositions
    var precompsFolderName = "Precomps"; // Name of the folder for new precompositions

    var data = readTSV(tsvPath);

    // Ensure the project has items
    if (app.project.items.length === 0) {
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
      throw new Error(
        "No composition found in the project. Please ensure your project has at least one composition."
      );
    }

    // Create a new folder for the generated compositions
    var outputFolder = app.project.items.addFolder(outputFolderName);

    // Create a nested folder for the precompositions within the output folder
    var precompsFolder = outputFolder.items.addFolder(precompsFolderName);

    for (var i = 0; i < data.length; i++) {
      createCompFromData(
        data[i],
        i + 1,
        mainComp,
        outputFolder,
        precompsFolder
      );
    }
  }

  main();
}
