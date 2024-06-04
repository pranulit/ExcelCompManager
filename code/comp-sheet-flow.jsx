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
    compsFromName();
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
  var button2 = myPanel.add("button", undefined, "Create a CSV / TXT");
  button2.preferredSize.width = checkboxesRowWidth;
  button2.onClick = function () {
    textPathNameExtract();
  };

  // Button for 'Content Import'
  var button3 = myPanel.add("button", undefined, "Import a CSV / TXT");
  button3.preferredSize.width = checkboxesRowWidth;
  button3.onClick = function () {
    compsFromSheet();
  };

  return myPanel;
}

function applySyntax(applyToText, applyToFiles) {
  var project = app.project;

  // Begin the undo group
  app.beginUndoGroup("Apply Syntax");

  function processComp(comp) {
    for (var j = 1; j <= comp.numLayers; j++) {
      var layer = comp.layer(j);

      if (applyToText && layer instanceof TextLayer) {
        if (layer.name.charAt(0) !== "@") {
          layer.name = "@" + layer.name;
        }
      }

      if (applyToFiles && layer.source instanceof FootageItem) {
        if (layer.name.charAt(0) !== "$" && layer.name.charAt(0) !== "#") {
          layer.name = "$" + layer.name;
        }
      }

      // Recursively process precompositions
      if (layer.source instanceof CompItem) {
        processComp(layer.source);
      }
    }
  }

  for (var i = 1; i <= project.numItems; i++) {
    if (project.item(i) instanceof CompItem) {
      processComp(project.item(i));
    }
  }

  // End the undo group
  app.endUndoGroup();
}

// Example usage
applySyntax(true, true);

// Function definitions
function compsFromName() {
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
      var lines = compsInput.split("\n");

      for (var i = 0; i < lines.length; i++) {
        lines[i] = lines[i].replace(/^\s+|\s+$/g, "");
      }

      // Manual filtering of empty lines
      var filteredLines = [];
      for (var i = 0; i < lines.length; i++) {
        if (lines[i] !== "") {
          filteredLines.push(lines[i]);
        }
      }
      lines = filteredLines;

      var files = filesInput.split("\n");

      for (var i = 0; i < files.length; i++) {
        files[i] = files[i].replace(/^\s+|\s+$/g, "");
      }

      // Manual filtering of empty file paths
      var filteredFiles = [];
      for (var i = 0; i < files.length; i++) {
        if (files[i] !== "") {
          filteredFiles.push(files[i]);
        }
      }
      files = filteredFiles;

      if (lines.length === 0) {
        return;
      }

      try {
        var mainFolder = app.project.items.addFolder("01_main");

        var importedFolder = app.project.items.addFolder("02_imported files");
      } catch (folderError) {
        alert("Error creating folders: " + folderError.message);
        return;
      }

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
          var parts = line.split("_");

          var duration = parts[2] ? parts[2].match(/(\d+)s/) : null;
          var formatMatch = parts[2] ? parts[2].match(/(\d+x\d+)/) : null;
          var format = formatMatch ? formatMatch[1] : defaultFormat;

          if (!duration) {
            duration = ["", "30"];
          }
          if (!compSettings[format]) {
            format = defaultFormat;
          }

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
            "$offline",
            comp.width,
            comp.height,
            1
          );
          solid.name = "$offline";

          if (files[i]) {
            var filePath = files[i];
            if (filePath !== "") {
              try {
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
              } catch (error) {
                alert(
                  "Error importing file: " + filePath + "\n" + error.message
                );
              }
            }
          }
        }
      } catch (e) {
        alert("Error in createCompositions: " + e.toString());
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

function textPathNameExtract() {
  {
    function chooseFilePath() {
      var file = File.openDialog(
        "Select a CSV, TSV, or delimited TXT file",
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

    function getDelimiter(filePath) {
      var parts = filePath.split(".");
      var extension = parts[parts.length - 1].toLowerCase();
      var delimiter;

      if ($.os.indexOf("Windows") !== -1) {
        delimiter = ";";
      } else {
        delimiter = ",";
      }

      switch (extension) {
        case "csv":
          return delimiter;
        case "tsv":
          return "\t";
        case "txt":
          return "\t"; // Default delimiter for txt files (can be changed)
        default:
          return delimiter; // Default delimiter if unknown
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

    // Main function or appropriate place in the script to set the file path
    var filePath = chooseFilePath(); // Get the file path using the dialog box
    if (!filePath) {
      throw new Error("Script terminated: No file selected.");
    }

    // Function to read and parse the delimited text file
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

    // Function to correctly parse a CSV line, considering quoted fields
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

        // Replace layers with corresponding imported files if the layer name starts with "$"
        if (layerName.indexOf("$") === 0) {
          var columnName = layerName.substring(1); // Remove the "$" symbol
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
      var uniqueID = new Date().getTime(); // Generate a unique ID based on the current timestamp
      var newComp = mainComp.duplicate();
      newComp.name = rowData["Composition Name"] + "_" + uniqueID; // Append unique ID to the composition name

      var importedFiles = {};

      // Import files and store them in the dictionary
      for (var key in rowData) {
        if (key.indexOf("$") === 0) {
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
      var outputFilePath =
        outputFolder + pathSeparator + comp.name + "_" + uniqueID + ".mp4"; // Append unique ID to file name

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

      if (!renderOptions.addToRenderQueue) {
        return; // Exit if user did not choose to add to render queue
      }

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
            if (renderOptions.outputModule) {
              var renderQueueItem = addToRenderQueueAndVerify(
                newComp,
                renderOptions.outputModule,
                renderOptions.outputFolder
              );
              if (renderQueueItem.status == 3013) {
                alert(
                  "Render queue item status is 3013. Please check the output module and destination folder settings."
                );
              }
            } else {
              messages.push("No output module selected for: " + newComp.name);
            }
          } else {
            var message = "Composition not found: " + compName;
            alert(message);
            messages.push(message);
          }
        }
      }

      // Start rendering process with feedback
      renderAndProvideFeedback();
    }

    main();
  }
}

function compsFromSheet() {
  {
    function chooseFilePath() {
      var file = File.openDialog(
        "Select a CSV, TSV, or delimited TXT file",
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

    function getDelimiter(filePath) {
      var parts = filePath.split(".");
      var extension = parts[parts.length - 1].toLowerCase();
      var delimiter;

      if ($.os.indexOf("Windows") !== -1) {
        delimiter = ";";
      } else {
        delimiter = ",";
      }

      switch (extension) {
        case "csv":
          return delimiter;
        case "tsv":
          return "\t";
        case "txt":
          return "\t"; // Default delimiter for txt files (can be changed)
        default:
          return delimiter; // Default delimiter if unknown
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

    // Main function or appropriate place in the script to set the file path
    var filePath = chooseFilePath(); // Get the file path using the dialog box
    if (!filePath) {
      throw new Error("Script terminated: No file selected.");
    }

    // Function to read and parse the delimited text file
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

    // Function to correctly parse a CSV line, considering quoted fields
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

        // Replace layers with corresponding imported files if the layer name starts with "$"
        if (layerName.indexOf("$") === 0) {
          var columnName = layerName.substring(1); // Remove the "$" symbol
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
      var uniqueID = new Date().getTime(); // Generate a unique ID based on the current timestamp
      var newComp = mainComp.duplicate();
      newComp.name = rowData["Composition Name"] + "_" + uniqueID; // Append unique ID to the composition name

      var importedFiles = {};

      // Import files and store them in the dictionary
      for (var key in rowData) {
        if (key.indexOf("$") === 0) {
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
      var outputFilePath =
        outputFolder + pathSeparator + comp.name + "_" + uniqueID + ".mp4"; // Append unique ID to file name

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

      var namesInput = dialog.add("edittext", undefined, "", {
        multiline: true,
      });
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

      var compsToRename = [];

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
            compsToRename.push(newComp);
          } else {
            var message = "Composition not found: " + compName;
            alert(message);
            messages.push(message);
          }
        }
      }

      var renamedComps = showRenameCompsDialog(compsToRename);
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
}

var myToolsPanel = createUI(this);

if (myToolsPanel instanceof Window) {
  myToolsPanel.center();
  myToolsPanel.show();
} else {
  myToolsPanel.layout.layout(true);
}
