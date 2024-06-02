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

// Function to determine the delimiter based on the file extension and OS
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
      var headers = file.readln().split(delimiter); // Read header line
      while (!file.eof) {
        var line = file.readln();
        var columns = line.split(delimiter);
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
    var newComp = mainComp.duplicate();
    newComp.name = rowData["Composition Name"];

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
    updateLayers(newComp, rowData, precompsFolder, importFolder, importedFiles);

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
    var outputFilePath = outputFolder + pathSeparator + comp.name + ".mp4";

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
    if (renderOptions.addToRenderQueue) {
      renderAndProvideFeedback();
    }
  }

  main();
}
