{
  // Function to display a file dialog and allow the user to select a CSV file.
  // Returns the full path of the selected file, or null if no file is selected.
  function chooseFilePath() {
    var file = File.openDialog(
      "Select a CSV",
      "Delimited Text files:*.csv;*.tsv;*.txt",
      false
    );
    if (file) {
      return file.fsName; // Return the full path of the file
    } else {
      alert("No file was selected.");
      return null; // Return null if no file is selected
    }
  }

  // Main function or appropriate place in the script to set the file path
  var filePath = chooseFilePath(); // Get the file path using the dialog box
  if (!filePath) {
    throw new Error("Script terminated: No file selected.");
  }

  // Function to determine the delimiter used in the selected file based on its content.
  // Returns ";" if the file uses semicolons, "," if it uses commas, and "\t" for TSV and TXT files.
  function getDelimiter(filePath) {
    var parts = filePath.split(".");
    var extension = parts[parts.length - 1].toLowerCase();
    var delimiter;

    // Default delimiter based on OS
    if ($.os.indexOf("Windows") !== -1) {
      delimiter = ";"; // Default for CSV on Windows
    } else {
      delimiter = ","; // Default for CSV on macOS and others
    }

    switch (extension) {
      case "csv":
        // Detect delimiter based on the content of the file
        var fileContent = readFile(filePath);
        if (fileContent.indexOf(";") > -1) {
          return ";";
        } else if (fileContent.indexOf(",") > -1) {
          return ",";
        }
        return delimiter;
      case "tsv":
      case "txt":
        return "\t"; // Default delimiter for tsv and txt files
      default:
        return delimiter; // Default delimiter if unknown
    }
  }

  // Function to read the content of the file
  function readFile(filePath) {
    var file = File(filePath);
    file.open("r");
    var content = file.read();
    file.close();
    return content;
  }

  // Function to display a dialog for selecting render settings and adding to the render queue.
  // Returns an object with the user's choices: whether to add to the render queue, the output module, and the output folder.
  function showRenderQueueDialog() {
    var dialog = new Window("dialog", "Render Queue Options");

    // Output module dropdown
    dialog.add("statictext", undefined, "Select Output Module:");
    var outputModuleDropdown = dialog.add("dropdownlist", undefined, []);

    // Create a temporary composition to access output module templates
    var tempComp = app.project.items.addComp("tempComp", 1920, 1080, 1, 1, 30);
    var renderQueueItem = app.project.renderQueue.items.add(tempComp);
    var outputModuleTemplates = renderQueueItem.outputModule(1).templates;

    for (var i = 0; i < outputModuleTemplates.length; i++) {
      if (outputModuleTemplates[i].indexOf("_HIDDEN") !== 0) {
        outputModuleDropdown.add("item", outputModuleTemplates[i]);
      }
    }

    // Remove temporary composition and render queue item
    renderQueueItem.remove();
    tempComp.remove();

    outputModuleDropdown.selection = 0; // Default selection

    // Folder selection for render destination
    dialog.add("statictext", undefined, "Select Destination Folder:");
    var folderGroup = dialog.add("group");
    folderGroup.orientation = "row";
    var folderPathText = folderGroup.add(
      "edittext",
      undefined,
      "~/Desktop/Render/"
    );
    folderPathText.characters = 50;
    var browseButton = folderGroup.add("button", undefined, "Browse");

    browseButton.onClick = function () {
      var folder = Folder.selectDialog("Select Render Destination Folder");
      if (folder) {
        folderPathText.text = folder.fsName;
      }
    };

    // Yes and No buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var yesButton = buttonGroup.add("button", undefined, "Yes");
    var noButton = buttonGroup.add("button", undefined, "No");

    var result = {
      addToRenderQueue: false,
      outputModule: null,
      outputFolder: null,
    };
    yesButton.onClick = function () {
      result.addToRenderQueue = true;
      result.outputModule = outputModuleDropdown.selection
        ? outputModuleDropdown.selection.text
        : null;
      result.outputFolder = folderPathText.text;
      dialog.close();
    };
    noButton.onClick = function () {
      result.addToRenderQueue = false;
      dialog.close();
    };

    dialog.show();
    return result;
  }

  // Function to read and parse a delimited text file (CSV, TSV, etc.).
  // Returns an array of objects, each representing a row in the file.
  function parseDocument(filePath) {
    var file = new File(filePath);
    var delimiter = getDelimiter(filePath);
    var data = [];
    if (file.open("r")) {
      var headers = parseLine(file.readln(), delimiter); // Read header line
      while (!file.eof) {
        var line = file.readln();
        var columns = parseLine(line, delimiter);
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

  // Function to correctly parse a line of a CSV file, considering quoted fields.
  // Returns an array of field values.
  function parseLine(line, delimiter) {
    var result = [];
    var insideQuote = false;
    var field = "";
    for (var i = 0; i < line.length; i++) {
      var currentChar = line.charAt(i);
      if (currentChar === '"') {
        insideQuote = !insideQuote; // Toggle the insideQuote flag
      } else if (currentChar === delimiter && !insideQuote) {
        result.push(field);
        field = "";
      } else {
        field += currentChar;
      }
    }
    result.push(field); // Add the last field
    return result;
  }

  // Function to import a file (e.g., an image or video) into the project and place it in the specified folder.
  // Returns the imported file or null if the file doesn't exist.
  function importFile(filePath, importFolder) {
    var file = new File(filePath);
    if (file.exists) {
      var importOptions = new ImportOptions(file);
      if (importOptions.canImportAs(ImportAsType.FOOTAGE)) {
        importOptions.importAs = ImportAsType.FOOTAGE;
      }
      var importedFile = app.project.importFile(importOptions);
      importedFile.parentFolder = importFolder; // Move the imported file to the specified folder
      return importedFile;
    } else {
      alert("File not found: " + filePath);
      return null;
    }
  }

  // Function to find a project item (e.g., composition, footage) by its name.
  // Returns the found project item or null if not found.
  function findProjectItemByName(name) {
    for (var i = 1; i <= app.project.numItems; i++) {
      if (app.project.item(i).name === name) {
        return app.project.item(i);
      }
    }
    return null;
  }

  // Function to show a dialog where the user can enter a suffix for precompositions.
  // Returns the entered suffix or an empty string if the dialog is canceled.
  function getPrecompSuffix() {
    var dialog = new Window("dialog", "Enter Suffix for Precomps");
    dialog.add(
      "statictext",
      undefined,
      "Enter the suffix to append to precomp names:"
    );
    var input = dialog.add("edittext", undefined, "");
    input.characters = 20;

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");

    okButton.onClick = function () {
      dialog.close(1);
    };

    cancelButton.onClick = function () {
      dialog.close(0);
    };

    if (dialog.show() === 1) {
      // Ensure that suffix is trimmed and check if it is empty
      var suffix = input.text.replace(/^\s+|\s+$/g, "");
      return suffix || ""; // If the suffix is empty, return an empty string
    } else {
      return ""; // If dialog was cancelled, return an empty string
    }
  }

  function containsLayerWithSymbols(comp, symbols) {
    for (var i = 1; i <= comp.layers.length; i++) {
      var layer = comp.layer(i);
      for (var j = 0; j < symbols.length; j++) {
        if (layer.name.indexOf(symbols[j]) !== -1) {
          return true; // Found a layer with one of the symbols
        }
      }
    }
    return false; // No layers with the symbols found
  }

  function updateLayers(
    comp,
    rowData,
    precompsFolder,
    importFolder,
    importedFiles,
    precompMap,
    suffix
  ) {
    if (!importedFiles) {
      alert("importedFiles is undefined!");
      return;
    }

    var symbols = ["@", "$", "#"];

    // Check if the composition has already been processed
    if (precompMap[comp.name]) {
      return;
    }

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

      // Replace layers with corresponding imported files if the layer name starts with "$"
      if (layerName.indexOf("$") === 0) {
        var columnName = layerName.substring(1);
        var importedFile = importedFiles[columnName];

        if (importedFile) {
          var originalStartTime = layer.startTime;
          var originalInPoint = layer.inPoint;
          var originalOutPoint = layer.outPoint;
          var originalDuration = layer.outPoint - layer.inPoint;
          var originalStretch = layer.stretch;
          var originalEnabled = layer.enabled;

          layer.replaceSource(importedFile, false);

          layer.startTime = originalStartTime;
          layer.inPoint = originalInPoint;
          layer.outPoint = originalOutPoint;
          layer.stretch = originalStretch;
          layer.enabled = originalEnabled;
        }
      }

      // Replace layers with corresponding project items if the layer name starts with "#"
      if (layerName.indexOf("#") === 0) {
        var columnName = layerName.substring(1);
        var projectItemName = rowData["#" + columnName];
        var projectItem = findProjectItemByName(projectItemName);
        if (projectItem) {
          var originalStartTime = layer.startTime;
          var originalInPoint = layer.inPoint;
          var originalOutPoint = layer.outPoint;
          var originalDuration = layer.outPoint - layer.inPoint;
          var originalStretch = layer.stretch;
          var originalEnabled = layer.enabled;

          layer.replaceSource(projectItem, false);

          layer.startTime = originalStartTime;
          layer.inPoint = originalInPoint;
          layer.outPoint = originalOutPoint;
          layer.stretch = originalStretch;
          layer.enabled = originalEnabled;
        }
      }

      // Recursively update and duplicate text layers in precompositions only if they have symbols
      if (layer.source instanceof CompItem) {
        var preComp = layer.source;

        if (containsLayerWithSymbols(preComp, symbols)) {
          if (!precompMap[preComp.name]) {
            var preCompDuplicate = preComp.duplicate();
            preCompDuplicate.name = preComp.name + suffix;
            preCompDuplicate.parentFolder = precompsFolder;

            // Mark the original precomp as processed to avoid duplication
            precompMap[preComp.name] = preCompDuplicate;

            // Update layers in the duplicated precomp
            updateLayers(
              preCompDuplicate,
              rowData,
              precompsFolder,
              importFolder,
              importedFiles,
              precompMap,
              suffix
            );
          }

          // Replace the layer with the duplicated precomp
          layer.replaceSource(precompMap[preComp.name], false);
          layer.enabled = true;
        }
      }
    }

    // Mark the top-level composition as processed to avoid reprocessing
    precompMap[comp.name] = comp;
  }

  // Function to update layers in a composition with CSV data and manage precompositions
  function updateLayersWithCSVData(
    comp,
    rowData,
    precompsFolder,
    importFolder,
    importedFiles,
    precompMap,
    suffix
  ) {
    if (!importedFiles) {
      alert("importedFiles is undefined!");
      return;
    }

    // Collect all affected layers and their precomps
    var affectedPrecomps = new Array();
    var allLayers = [];

    for (var i = 1; i <= comp.layers.length; i++) {
      var layer = comp.layer(i);
      var layerName = layer.name;

      // Update text layers based on rowData
      if (rowData[layerName]) {
        var textValue = rowData[layerName];
        if (textValue && layer.property("Source Text") != null) {
          layer.property("Source Text").setValue(textValue);
        }
      }

      // Collect layers for processing
      allLayers.push(layer);

      // Identify precomps used by layers affected by CSV data
      if (layer.source instanceof CompItem) {
        var precompName = layer.source.name;
        if (rowData[layerName]) {
          // Only consider if layer is affected by CSV data
          if (affectedPrecomps.indexOf(precompName) === -1) {
            affectedPrecomps.push(precompName);
          }
        }
      }
    }

    // Duplicate only the affected precomps
    for (var i = 0; i < affectedPrecomps.length; i++) {
      var precompName = affectedPrecomps[i];
      if (!precompMap[precompName]) {
        var preComp = findProjectItemByName(precompName);
        if (preComp && preComp instanceof CompItem) {
          // Duplicate and name the precomp
          var preCompDuplicate = preComp.duplicate();
          preCompDuplicate.name = precompName + suffix; // Ensure the name includes suffix
          preCompDuplicate.parentFolder = precompsFolder;

          // Add the newly duplicated precomp to the map
          precompMap[precompName] = preCompDuplicate;
        }
      }
    }

    // Replace layers' sources with the duplicated or existing precomps
    for (var i = 0; i < allLayers.length; i++) {
      var layer = allLayers[i];
      if (layer.source instanceof CompItem) {
        var precompName = layer.source.name;
        var newPrecomp = precompMap[precompName];
        if (newPrecomp) {
          layer.replaceSource(newPrecomp, false);
          layer.enabled = true; // Ensure the layer is visible
        } else {
          alert("Precomp not found in precompMap: " + precompName);
        }
      }
    }
  }

  // Function to create a new composition based on rowData and manage its import and layer updates
  function createCompFromData(
    rowData,
    mainComp,
    outputFolder,
    precompsFolder,
    importFolder,
    existingCompNames,
    precompMap,
    suffix
  ) {
    var newName = rowData["Composition Name"] + "_NEW";

    //   var baseName = rowData["Composition Name"] + "_NEW_v";
    //   var newName = baseName + "001";
    //   var count = 1;

    // var count = 1;
    // var newName = baseName + ("000" + count).slice(-3);
    // var isUnique = false;

    // // Loop until a unique name is found
    // while (!isUnique) {
    //     isUnique = true; // Assume it's unique initially

    //     // Manually check if newName exists in the array
    //     for (var i = 0; i < existingCompNames.length; i++) {
    //         if (existingCompNames[i] === newName) {
    //             isUnique = false; // Found a match, name is not unique
    //             count++;
    //             newName = baseName + ("000" + count).slice(-3);
    //             break; // Exit the loop early since we found a match
    //         }
    //     }
    // }

    // newName now holds a unique composition name

    existingCompNames.push(newName);

    var newComp = mainComp.duplicate();
    newComp.name = newName;

    var importedFiles = {};

    // Import files and store them in the dictionary
    for (var key in rowData) {
      if (rowData.hasOwnProperty(key) && key.indexOf("$") === 0) {
        var filePath = rowData[key];
        if (filePath) {
          var importedFile = importFile(filePath, importFolder);
          if (importedFile) {
            importedFiles[key.substring(1)] = importedFile;
          }
        }
      }
    }

    // Call updateLayers to update layers in the duplicated composition and its precompositions
    updateLayers(
      newComp,
      rowData,
      precompsFolder,
      importFolder,
      importedFiles,
      precompMap,
      suffix
    );

    // Move the new composition to the specified folder
    newComp.parentFolder = outputFolder;

    return newComp;
  }

  function addToRenderQueueAndVerify(comp, outputModule, outputFolder) {
    // Ensure the output folder exists
    var outputDir = new Folder(outputFolder);
    if (!outputDir.exists) {
      var created = outputDir.create();
      if (!created) {
        alert("Failed to create output directory: " + outputFolder);
        return null;
      }
    }

    var renderQueueItem = app.project.renderQueue.items.add(comp);
    renderQueueItem.outputModule(1).applyTemplate(outputModule);

    // Use appropriate path separator
    var pathSeparator = Folder.fs == "Macintosh" ? "/" : "\\";
    var uniqueID = new Date().getTime(); // Generate a unique ID based on the current timestamp
    var outputFilePath = outputFolder + pathSeparator + comp.name; // Append unique ID to file name

    renderQueueItem.outputModule(1).file = new File(outputFilePath);

    return renderQueueItem;
  }

  function checkRenderStatus(renderQueueItem) {
    var status = renderQueueItem.status;
    var statusMessage = "";

    switch (status) {
      case RQItemStatus.QUEUED:
        statusMessage = "Queued";
        break;
      case RQItemStatus.RENDERING:
        statusMessage = "Rendering";
        break;
      case RQItemStatus.DONE:
        statusMessage = "Done";
        break;
      case RQItemStatus.ERR_STOPPED:
        statusMessage = "Error Stopped";
        break;
      case RQItemStatus.USER_STOPPED:
        statusMessage = "User Stopped";
        break;
      default:
        statusMessage = "Unknown Status (" + status + ")";
        break;
    }

    return statusMessage;
  }

  function renderAndProvideFeedback() {
    var renderQueue = app.project.renderQueue;
    var totalItems = renderQueue.numItems;
    var completedItems = 0;

    for (var i = 1; i <= totalItems; i++) {
      var renderQueueItem = renderQueue.item(i);

      if (renderQueueItem.status == RQItemStatus.QUEUED) {
        renderQueue.render();
      }

      while (renderQueueItem.status == RQItemStatus.RENDERING) {
        $.sleep(1000); // Wait for 1 second
      }

      completedItems++;
    }

    alert("Rendering complete. " + completedItems + " items rendered.");
  }

  function showRenameCompsDialog(comps) {
    var dialog = new Window("dialog", "Rename Compositions");

    dialog.add("statictext", undefined, "Enter new names (one per line):");

    var previousNamesGroup = dialog.add("group");
    previousNamesGroup.orientation = "column";
    for (var i = 0; i < comps.length; i++) {
      previousNamesGroup.add(
        "statictext",
        undefined,
        "Current Name: " + comps[i].name
      );
    }

    var namesInput = dialog.add("edittext", undefined, "", { multiline: true });
    namesInput.preferredSize = { width: 300, height: 200 };

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");

    okButton.onClick = function () {
      dialog.close(1);
    };
    cancelButton.onClick = function () {
      dialog.close(0);
    };

    if (dialog.show() === 1) {
      var newNames = namesInput.text.split("\n");
      for (var i = 0; i < newNames.length; i++) {
        newNames[i] = newNames[i].replace(/^\s+|\s+$/g, ""); // Trim whitespace
      }
      var renamedComps = [];
      for (var i = 0; i < comps.length; i++) {
        renamedComps.push({
          oldName: comps[i].name,
          newName: newNames[i] || comps[i].name, // Retain the default name if not renamed
        });
      }
      return renamedComps;
    } else {
      return null; // User canceled the renaming
    }
  }

  function showNamingOptionDialog() {
    var dialog = new Window("dialog", "New Composition Naming");

    dialog.add("statictext", undefined, "Choose naming option:");
    var radioGroup = dialog.add("group");
    radioGroup.orientation = "column";

    var manualNamingOption = radioGroup.add(
      "radiobutton",
      undefined,
      "Insert new comp naming"
    );
    var spreadsheetNamingOption = radioGroup.add(
      "radiobutton",
      undefined,
      "Comp naming is in the spreadsheet"
    );

    manualNamingOption.value = true; // Set default option

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var okButton = buttonGroup.add("button", undefined, "OK");
    var cancelButton = buttonGroup.add("button", undefined, "Cancel");

    var result = { useSpreadsheetNaming: false };

    okButton.onClick = function () {
      result.useSpreadsheetNaming = spreadsheetNamingOption.value;
      dialog.close();
    };
    cancelButton.onClick = function () {
      dialog.close();
    };

    dialog.show();
    return result;
  }

  // Main function to handle the overall process, including user input, composition creation, and rendering
  function main() {
    var suffix = getPrecompSuffix();

    var outputFolderName = "00_Generated Comps";
    var precompsFolderName = "Precomps";
    var importFolderName = "Imported Files";

    var data = parseDocument(filePath);
    var messages = [];
    var existingCompNames = [];
    var precompMap = {}; // Add this to keep track of precomps
    var templateCompositions = []; // Initialize the array to store found compositions

    if (app.project.items.length === 0) {
      var message =
        "No items in the project. Please ensure your project has at least one composition.";
      alert(message);
      throw new Error(message);
    }

    var outputFolder = app.project.items.addFolder(outputFolderName);
    var precompsFolder = outputFolder.items.addFolder(precompsFolderName);
    var importFolder = outputFolder.items.addFolder(importFolderName);

    var compsToRename = [];

    var namingOption = showNamingOptionDialog();

    for (var i = 0; i < data.length; i++) {
      var compName = data[i]["Composition Name"];
      if (compName) {
        var mainComp = null;
        for (var j = 1; j <= app.project.items.length; j++) {
          if (
            app.project.items[j] instanceof CompItem &&
            app.project.items[j].name === compName
          ) {
            mainComp = app.project.items[j];
            break;
          }
        }

        if (mainComp) {
          // Store the found composition in templateCompositions
          templateCompositions.push(mainComp);

          var newComp = createCompFromData(
            data[i],
            mainComp,
            outputFolder,
            precompsFolder,
            importFolder,
            existingCompNames,
            precompMap, // Pass the precompMap
            suffix
          );
          compsToRename.push(newComp);
        } else {
          var message = "Composition not found: " + compName;
          alert(message);
          messages.push(message);
        }
      }
    }

    var renamedComps = [];
    if (namingOption.useSpreadsheetNaming) {
      for (var i = 0; i < compsToRename.length; i++) {
        var newName = data[i]["New Comp Name"];
        renamedComps.push({
          oldName: compsToRename[i].name,
          newName: newName || compsToRename[i].name, // Retain the default name if not provided
        });
      }
    } else {
      renamedComps = showRenameCompsDialog(compsToRename);
    }

    if (renamedComps) {
      for (var i = 0; i < renamedComps.length; i++) {
        var compInfo = renamedComps[i];
        var comp = findProjectItemByName(compInfo.oldName);
        if (comp) {
          comp.name = compInfo.newName;
        }
      }
    }

    // Show dialog to decide whether to add to render queue
    var renderOptions = showRenderQueueDialog();

    if (!renderOptions.addToRenderQueue) {
      return; // Exit if user did not choose to add to render queue
    }

    if (renamedComps) {
      for (var i = 0; i < renamedComps.length; i++) {
        var compInfo = renamedComps[i];
        var comp = findProjectItemByName(compInfo.newName);
        if (comp && renderOptions.outputModule) {
          var renderQueueItem = addToRenderQueueAndVerify(
            comp,
            renderOptions.outputModule,
            renderOptions.outputFolder
          );
          if (renderQueueItem.status == 3013) {
            alert(
              "Render queue item status is 3013. Please check the output module and destination folder settings."
            );
          }
        }
      }
    }

    // Start rendering process with feedback
    renderAndProvideFeedback();
  }

  main();
}
