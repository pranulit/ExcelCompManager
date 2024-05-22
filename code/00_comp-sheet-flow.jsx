function createUI(thisObj) {
  var myPanel =
    thisObj instanceof Panel
      ? thisObj
      : new Window("palette", "Comp-Sheet-Flow", undefined, {
          resizeable: true,
        });
  myPanel.orientation = "column";

  // Button for 'Comps from Names'
  var button1 = myPanel.add("button", undefined, "Comps from Names");
  button1.onClick = function () {
    compsFromNames();
  };

  // Group for the new button and checkboxes
  var newButtonGroup = myPanel.add("group");
  newButtonGroup.orientation = "row";

  // Checkboxes
  var textCheckbox = newButtonGroup.add("checkbox", undefined, "Text");
  var filesCheckbox = newButtonGroup.add("checkbox", undefined, "Files");

  // Button for applying syntax
  var applySyntaxButton = newButtonGroup.add(
    "button",
    undefined,
    "Apply Syntax"
  );
  applySyntaxButton.onClick = function () {
    applySyntax(textCheckbox.value, filesCheckbox.value);
  };

  // Calculate width of the checkboxes row
  var checkboxesRowWidth =
    textCheckbox.preferredSize.width +
    filesCheckbox.preferredSize.width +
    applySyntaxButton.preferredSize.width;

  // Set width of the first button
  button1.preferredSize.width = checkboxesRowWidth;

  // Button for 'Content Extract'
  var button2 = myPanel.add("button", undefined, "Create a CSV/TXT");
  button2.preferredSize.width = checkboxesRowWidth;
  button2.onClick = function () {
    contentExtract();
  };

  // Button for 'Content Import'
  var button3 = myPanel.add("button", undefined, "Import a TXT");
  button3.preferredSize.width = checkboxesRowWidth;
  button3.onClick = function () {
    contentImport();
  };

  return myPanel;
}

function applySyntax(applyToText, applyToFiles) {
  var project = app.project;

  for (var i = 1; i <= project.numItems; i++) {
    if (project.item(i) instanceof CompItem) {
      var comp = project.item(i);

      for (var j = 1; j <= comp.numLayers; j++) {
        var layer = comp.layer(j);

        if (applyToText && layer instanceof TextLayer) {
          if (layer.name.charAt(0) !== "^") {
            layer.name = "^" + layer.name;
          }
        }

        if (applyToFiles && layer.source instanceof FootageItem) {
          if (layer.name.charAt(0) !== ">") {
            layer.name = ">" + layer.name;
          }
        }
      }
    }
  }
}

// Function definitions
function compsFromNames() {
  (function () {
    var myPanel =
      this instanceof Panel
        ? this
        : new Window("palette", "Create Compositions", undefined, {
            resizeable: true,
          });

    if (myPanel != null) {
      var textGroup = myPanel.add("group");
      textGroup.orientation = "row";
      textGroup.add("statictext", undefined, "Enter Composition Names:");

      var inputGroup = myPanel.add("group");
      inputGroup.orientation = "row";
      var compsInput = inputGroup.add("edittext", undefined, "", {
        multiline: true,
      });
      compsInput.size = [300, 100];

      var fileGroup = myPanel.add("group");
      fileGroup.orientation = "row";
      fileGroup.add(
        "statictext",
        undefined,
        "Enter Offline File Paths (optional):"
      );

      var fileInputGroup = myPanel.add("group");
      fileInputGroup.orientation = "row";
      var filesInput = fileInputGroup.add("edittext", undefined, "", {
        multiline: true,
      });
      filesInput.size = [300, 100];

      var browseFilesButton = fileGroup.add("button", undefined, "Browse...");
      browseFilesButton.onClick = function () {
        var filePaths = File.openDialog(
          "Select files for offline file paths",
          "",
          true
        );
        if (filePaths) {
          var pathArray = [];
          for (var i = 0; i < filePaths.length; i++) {
            pathArray.push(filePaths[i].fsName);
          }
          filesInput.text = pathArray.join("\n");
        }
      };

      var fpsGroup = myPanel.add("group");
      fpsGroup.orientation = "row";
      fpsGroup.add("statictext", undefined, "Select FPS:");
      var fpsDropdown = fpsGroup.add("dropdownlist", undefined, [
        "24",
        "25",
        "30",
        "60",
      ]);
      fpsDropdown.selection = 1;

      var buttonGroup = myPanel.add("group");
      buttonGroup.orientation = "row";
      var createBtn = buttonGroup.add(
        "button",
        undefined,
        "Create Compositions"
      );
      createBtn.onClick = function () {
        createCompositions(
          compsInput.text,
          filesInput.text,
          parseFloat(fpsDropdown.selection.text)
        );
      };
    }

    function createCompositions(compsInput, filesInput, fps) {
      alert("createCompositions called");

      var lines = compsInput.split("\n");
      for (var i = 0; i < lines.length; i++) {
        lines[i] = lines[i].replace(/^\s+|\s+$/g, "");
      }
      lines = lines.filter(function (line) {
        return line !== "";
      });

      var files = filesInput.split("\n");
      for (var i = 0; i < files.length; i++) {
        files[i] = files[i].replace(/^\s+|\s+$/g, "");
      }
      files = files.filter(function (line) {
        return line !== "";
      });

      alert("Compositions Input: " + lines.join(", "));
      alert("Files Input: " + files.join(", "));
      alert("FPS: " + fps);

      if (lines.length === 0) {
        alert("No composition names provided.");
        return;
      }

      var mainFolder = app.project.items.addFolder("01_main");
      var importedFolder = app.project.items.addFolder("02_imported files");

      var compSettings = {
        "16x9": { width: 1920, height: 1080 },
        "1x1": { width: 1080, height: 1080 },
        "4x5": { width: 1080, height: 1350 },
        "9x16": { width: 1080, height: 1920 },
      };
      var defaultFormat = "16x9";

      app.beginUndoGroup("Create Compositions");

      try {
        for (var i = 0; i < lines.length; i++) {
          var line = lines[i];
          alert("Processing line: " + line);
          var parts = line.split("_");
          if (parts.length < 3) {
            alert("Invalid composition name format: " + line);
            continue;
          }

          var duration = parts[2].match(/(\d+)s/);
          var formatMatch = parts[2].match(/(\d+x\d+)/);
          var format = formatMatch ? formatMatch[1] : defaultFormat;

          if (duration && compSettings[format]) {
            var comp = app.project.items.addComp(
              parts[0],
              compSettings[format].width,
              compSettings[format].height,
              1,
              parseFloat(duration[1]),
              fps
            );
            comp.parentFolder = mainFolder;
            comp.name = line;

            var solid = comp.layers.addSolid(
              [1, 1, 1],
              ">offline",
              comp.width,
              comp.height,
              1
            );
            solid.name = ">offline";

            if (files[i]) {
              var filePath = files[i];
              alert("Processing file: " + filePath);
              if (filePath !== "") {
                var importOptions = new ImportOptions(File(filePath));
                if (importOptions.canImportAs(ImportAsType.FOOTAGE)) {
                  var footage = app.project.importFile(importOptions);
                  footage.parentFolder = importedFolder;
                  var newLayer = comp.layers.add(footage);
                  newLayer.moveBefore(solid);
                  solid.remove();
                } else {
                  alert(
                    "File at " + filePath + " cannot be imported as footage."
                  );
                }
              }
            }
          } else {
            alert("Could not determine valid duration/format for: " + line);
          }
        }
      } catch (e) {
        alert("Error: " + e.toString());
      }

      app.endUndoGroup();
    }

    if (myPanel instanceof Window) {
      myPanel.center();
      myPanel.show();
    } else {
      myPanel.layout.layout(true);
    }
  })();
}

function contentExtract() {
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
  // Define a function to create a user interface window in After Effects
  function createUI() {
    // Create a new dialog window titled "Select Compositions"
    var dlg = new Window("dialog", "Select Compositions", undefined, {
      resizeable: false,
    });
    dlg.size = [400, 400];

    // Add a group for the search bar
    var searchGroup = dlg.add("group");
    searchGroup.orientation = "row";
    // Add a text edit field and a search button to the group
    var searchEdit = searchGroup.add("edittext", undefined, "");
    searchEdit.characters = 20;
    var searchButton = searchGroup.add("button", undefined, "Search");

    // Add a list box to display items with multi-selection enabled
    var listBox = dlg.add("listbox", [0, 0, 380, 300], [], {
      multiselect: true,
    });
    listBox.alignment = "fill";

    // Populate the list box initially
    populateListBox(listBox, ""); // Assuming a function populateListBox is defined

    // Define the search button click behavior
    searchButton.onClick = function () {
      populateListBox(listBox, searchEdit.text);
    };

    // Add a group for buttons
    var buttons = dlg.add("group");
    buttons.orientation = "row";
    buttons.alignment = "right";
    var cancelButton = buttons.add("button", undefined, "Cancel", {
      name: "cancel",
    });
    var saveCSVButton = buttons.add("button", undefined, "Save as CSV", {
      name: "csv",
    });
    var saveTXTButton = buttons.add("button", undefined, "Save as TXT", {
      name: "txt",
    });

    // Define the cancel button behavior
    cancelButton.onClick = function () {
      dlg.close();
    };

    // Define the save button behavior for CSV
    saveCSVButton.onClick = function () {
      saveCompositions(listBox, "CSV");
      dlg.close();
    };

    // Define the save button behavior for TXT
    saveTXTButton.onClick = function () {
      saveCompositions(listBox, "TXT");
      dlg.close();
    };

    // Lay out the dialog elements, center it, and display it
    dlg.layout.layout(true);
    dlg.center();
    dlg.show();
  }

  // Define a function to save the selected compositions
  function saveCompositions(listBox, format) {
    try {
      var selectedCompositions = getSelectedCompositions(listBox);
      alert("Selected " + selectedCompositions.length + " compositions.");
      if (selectedCompositions.length === 0) {
        alert("No compositions are selected or associated properly.");
        return;
      }
      exportSelectedCompositions(selectedCompositions, format);
    } catch (error) {
      alert("Error: " + error.toString());
    }
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
      if (
        compItem.name.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1
      ) {
        var item = listBox.add("item", compItem.name);
        item.compItem = compItem;
      }
    }
  }

  function searchPrecomps(comp, parentCompData, project) {
    if (!parentCompData.textLayers) {
      parentCompData.textLayers = {};
    }
    if (!parentCompData.fileLayers) {
      parentCompData.fileLayers = {};
    }
    if (!parentCompData.specialLayers) {
      parentCompData.specialLayers = {};
    }

    var hasGreaterThanLayer = false;

    for (var k = 1; k <= comp.numLayers; k++) {
      var layer = comp.layer(k);

      // Extract text from text layers
      if (layer instanceof TextLayer && layer.property("Source Text") != null) {
        var textSource = layer.property("Source Text").value;
        if (textSource) {
          var textContent = textSource.text.replace(/[\r\n]+/g, " ");
          parentCompData.textLayers[layer.name] = textContent;
        }
      }

      // Capture all layers that start with ">"
      if (layer.name.substring(0, 1) === ">") {
        hasGreaterThanLayer = true;
        var filePath =
          layer.source && layer.source.file
            ? layer.source.file.fsName
            : "no path retrieved";
        parentCompData.fileLayers[layer.name] = filePath;
      }

      // Check if layer name starts with "#"
      if (layer.name.charAt(0) === "#") {
        if (layer.source) {
          var sourceName = getSourceName(layer.source.name, project);
          if (sourceName) {
            parentCompData.specialLayers[layer.name] = sourceName;
          } else {
            parentCompData.specialLayers[layer.name] = "Source not found";
          }
        } else {
          parentCompData.specialLayers[layer.name] = "No source associated";
        }
      }
    }

    // Mark the composition with hasGreaterThanLayer flag
    parentCompData.hasGreaterThanLayer = hasGreaterThanLayer;
  }

  // Helper function to get source name from project
  function getSourceName(layerName, project) {
    for (var i = 1; i <= project.items.length; i++) {
      var item = project.items[i];
      if (item instanceof FootageItem && item.name === layerName) {
        return item.name;
      }
    }
    return null; // Return null if source name is not found
  }

  function exportSelectedCompositions(compCheckboxes, format) {
    var comps = [];
    var uniqueTextLayerNames = {};
    var uniqueFileLayerNames = {};
    var uniqueSpecialLayerNames = {};

    var project = app.project;

    for (var i = 0; i < compCheckboxes.length; i++) {
      var compItem = compCheckboxes[i];
      if (!compItem || !compItem.name || !compItem.numLayers) {
        continue;
      }

      var parentCompData = {
        textLayers: {},
        fileLayers: {},
        specialLayers: {},
        hasGreaterThanLayer: false,
      };
      searchPrecomps(compItem, parentCompData, project);

      var compObj = {
        name: compItem.name,
        textLayersContent: parentCompData.textLayers,
        fileLayersContent: parentCompData.fileLayers,
        specialLayersContent: parentCompData.specialLayers,
        hasGreaterThanLayer: parentCompData.hasGreaterThanLayer,
      };

      comps.push(compObj);

      for (var layerName in parentCompData.textLayers) {
        uniqueTextLayerNames[layerName] = true;
      }
      for (var layerName in parentCompData.fileLayers) {
        uniqueFileLayerNames[layerName] = true;
      }
      for (var layerName in parentCompData.specialLayers) {
        uniqueSpecialLayerNames[layerName] = true;
      }
    }

    var delimiter = format === "CSV" ? "," : "\t";
    generateAndSaveCSV(
      comps,
      uniqueTextLayerNames,
      uniqueFileLayerNames,
      uniqueSpecialLayerNames,
      delimiter
    );
  }

  function generateAndSaveCSV(
    comps,
    uniqueTextLayerNames,
    uniqueFileLayerNames,
    uniqueSpecialLayerNames,
    delimiter
  ) {
    var headers = ["Composition Name"];

    for (var layerName in uniqueTextLayerNames) {
      if (uniqueTextLayerNames.hasOwnProperty(layerName)) {
        headers.push(layerName);
      }
    }

    for (layerName in uniqueFileLayerNames) {
      if (uniqueFileLayerNames.hasOwnProperty(layerName)) {
        headers.push(layerName);
      }
    }

    for (layerName in uniqueSpecialLayerNames) {
      if (uniqueSpecialLayerNames.hasOwnProperty(layerName)) {
        headers.push(layerName);
      }
    }

    var csvContent = headers.join(delimiter) + "\n";

    for (var i = 0; i < comps.length; i++) {
      var compObj = comps[i];
      var compRow = ['"' + compObj.name.replace(/"/g, '""') + '"'];

      for (var j = 1; j < headers.length; j++) {
        var header = headers[j];
        var actualLayerName = header;
        var content = "";

        if (compObj.textLayersContent.hasOwnProperty(actualLayerName)) {
          content = compObj.textLayersContent[actualLayerName];
        } else if (compObj.fileLayersContent.hasOwnProperty(actualLayerName)) {
          if (compObj.hasGreaterThanLayer) {
            content = compObj.fileLayersContent[actualLayerName];
          } else {
            content = "";
          }
        } else if (
          compObj.specialLayersContent.hasOwnProperty(actualLayerName)
        ) {
          content = compObj.specialLayersContent[actualLayerName];
        }

        compRow.push('"' + (content ? content.replace(/"/g, '""') : "") + '"');
      }

      csvContent += compRow.join(delimiter) + "\n";
    }

    saveDelimitedFile(csvContent, delimiter);
  }

  function saveDelimitedFile(content, delimiter) {
    var fileExtension = delimiter === "," ? "*.csv" : "*.txt";
    var file = new File(File.saveDialog("Save your file", fileExtension));
    if (file) {
      file.encoding = "UTF-8";
      file.open("w");
      if (file.write(content)) {
        alert("File saved successfully!");
      } else {
        alert("Failed to write to file.");
      }
      file.close();
    } else {
      alert("File save cancelled or failed to open file.");
    }
  }

  createUI();
  //works
}

function contentImport() {
  {
    function chooseFilePath() {
      var file = File.openDialog(
        "Select a TSV file",
        "TSV files:*.tsv;*.txt",
        false
      );
      if (file) {
        return file.fsName; // Return the full path of the file
      } else {
        alert("No file was selected.");
        return null; // Return null if no file is selected
      }
    }

    // Function to display UI for render settings and adding to render queue
    function showRenderQueueDialog() {
      var dialog = new Window("dialog", "Render Queue Options");

      // Output module dropdown
      dialog.add("statictext", undefined, "Select Output Module:");
      var outputModuleDropdown = dialog.add("dropdownlist", undefined, []);

      // Create a temporary composition to access output module templates
      var tempComp = app.project.items.addComp(
        "tempComp",
        1920,
        1080,
        1,
        1,
        30
      );
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

      // Yes and No buttons
      var buttonGroup = dialog.add("group");
      buttonGroup.orientation = "row";
      var yesButton = buttonGroup.add("button", undefined, "Yes");
      var noButton = buttonGroup.add("button", undefined, "No");

      var result = { addToRenderQueue: false, outputModule: null };
      yesButton.onClick = function () {
        result.addToRenderQueue = true;
        result.outputModule = outputModuleDropdown.selection
          ? outputModuleDropdown.selection.text
          : null;
        dialog.close();
      };
      noButton.onClick = function () {
        result.addToRenderQueue = false;
        dialog.close();
      };

      dialog.show();
      return result;
    }

    // Main function or appropriate place in the script to set the file path
    var filePath = chooseFilePath(); // Get the file path using the dialog box
    if (!filePath) {
      throw new Error("Script terminated: No file selected.");
    }

    // Function to read and parse the TSV file
    function parseDocument(filePath) {
      var file = new File(filePath);
      var data = [];
      if (file.open("r")) {
        var headers = file.readln().split("\t"); // Read header line
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
    function updateLayers(
      comp,
      rowData,
      precompsFolder,
      importFolder,
      importedFiles
    ) {
      if (!importedFiles) {
        alert("importedFiles is undefined!");
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

        // Replace layers with corresponding imported files if the layer name starts with ">"
        if (layerName.indexOf(">") === 0) {
          var columnName = layerName.substring(1); // Remove the ">" symbol
          var importedFile = importedFiles[columnName];

          if (importedFile) {
            layer.replaceSource(importedFile, false);
            layer.enabled = true; // Ensure the layer is visible
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
          }
        }

        // Recursively update and duplicate text layers in precompositions
        if (layer.source instanceof CompItem) {
          var preCompDuplicate = layer.source.duplicate();
          preCompDuplicate.name = comp.name + " - " + layerName;
          preCompDuplicate.parentFolder = precompsFolder;

          // Update layers in the duplicated precomp
          updateLayers(
            preCompDuplicate,
            rowData,
            precompsFolder,
            importFolder,
            importedFiles
          );

          // Update the layer to use the new precomp
          layer.replaceSource(preCompDuplicate, false);
          layer.enabled = true; // Ensure the layer is visible
        }
      }
    }

    // Function to create a comp from data
    function createCompFromData(
      rowData,
      mainComp,
      outputFolder,
      precompsFolder,
      importFolder
    ) {
      var newComp = mainComp.duplicate();
      newComp.name = rowData["Composition Name"];

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

      // Update layers in the duplicated composition and its precompositions
      updateLayers(
        newComp,
        rowData,
        precompsFolder,
        importFolder,
        importedFiles
      );

      // Move the new composition to the specified folder
      newComp.parentFolder = outputFolder;

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
        var message =
          "No items in the project. Please ensure your project has at least one composition.";
        alert(message);
        throw new Error(message);
      }

      // Create a new folder for the generated compositions
      var outputFolder = app.project.items.addFolder(outputFolderName);

      // Create a nested folder for the precompositions within the output folder
      var precompsFolder = outputFolder.items.addFolder(precompsFolderName);

      // Create a folder for imported files within the output folder
      var importFolder = outputFolder.items.addFolder(importFolderName);

      // Show dialog to decide whether to add to render queue
      var renderOptions = showRenderQueueDialog();

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
            var newComp = createCompFromData(
              data[i],
              mainComp,
              outputFolder,
              precompsFolder,
              importFolder
            );
            if (renderOptions.addToRenderQueue && renderOptions.outputModule) {
              var renderQueueItem = app.project.renderQueue.items.add(newComp);
              renderQueueItem
                .outputModule(1)
                .applyTemplate(renderOptions.outputModule);
            }
          } else {
            var message = "Composition not found: " + compName;
            alert(message);
            messages.push(message);
          }
        }
      }

      // Show all messages at the end
      var successMessage =
        "Script completed successfully.\n\n" + messages.join("\n");
      alert(successMessage);
    }

    main();
  }
}

var myToolsPanel = createUI(this);

if (myToolsPanel instanceof Window) {
  myToolsPanel.center();
  myToolsPanel.show();
} else {
  myToolsPanel.layout.layout(true);
}

//working
