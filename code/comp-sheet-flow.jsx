function createUI(thisObj) {
  var myPanel =
    thisObj instanceof Panel
      ? thisObj
      : new Window("palette", "Comp-Sheet-Flow", undefined, {
          resizeable: true,
        });
  myPanel.orientation = "column";

  // Button for applying/removing syntax
  var applyRemoveSyntaxButton = myPanel.add(
    "button",
    undefined,
    "Apply / Remove Syntax"
  );
  applyRemoveSyntaxButton.onClick = function () {
    openApplyRemoveSyntaxWindow();
  };

  // Button for 'Content Extract'
  var button2 = myPanel.add("button", undefined, "Create a CSV");
  button2.onClick = function () {
    textPathNameExtract();
  };

  // Button for 'Content Import'
  var button3 = myPanel.add("button", undefined, "Import a CSV");
  button3.onClick = function () {
    compsFromSheet();
  };

  return myPanel;
}

function openApplyRemoveSyntaxWindow() {
  // Change the dialog type to "palette" for a modeless window
  var dialog = new Window("palette", "Apply / Remove Syntax");
  dialog.orientation = "column";

  // Search bar for filtering compositions
  var searchGroup = dialog.add("group");
  searchGroup.orientation = "row";
  var searchLabel = searchGroup.add("statictext", undefined, "Search:");
  var searchInput = searchGroup.add("edittext", undefined, "");
  searchInput.characters = 20;
  var searchButton = searchGroup.add("button", undefined, "Search");

  // Listbox for selecting compositions, now with size 400x400
  var compsList = dialog.add("listbox", undefined, [], {
    multiselect: true,
  });
  compsList.preferredSize = [400, 400]; // Set the size to 400x400

  // Populate the listbox with composition names
  var project = app.project;
  function populateCompsList(filterText) {
    compsList.removeAll();
    for (var i = 1; i <= project.numItems; i++) {
      if (project.item(i) instanceof CompItem) {
        if (
          !filterText ||
          project
            .item(i)
            .name.toLowerCase()
            .indexOf(filterText.toLowerCase()) !== -1
        ) {
          compsList.add("item", project.item(i).name);
        }
      }
    }
  }

  // Initial population of the listbox
  populateCompsList("");

  // Event handler for the search button
  searchButton.onClick = function () {
    var searchText = searchInput.text;
    populateCompsList(searchText);
  };

  // Checkboxes for applying/removing syntax
  var optionsGroup = dialog.add("group");
  optionsGroup.orientation = "row";
  var textCheckbox = optionsGroup.add("checkbox", undefined, "@");
  var pathCheckbox = optionsGroup.add("checkbox", undefined, "$");
  var projectItemCheckbox = optionsGroup.add("checkbox", undefined, "#");

  // Buttons for applying or removing syntax
  var buttonGroup = dialog.add("group");
  buttonGroup.orientation = "row";
  var applyButton = buttonGroup.add("button", undefined, "Apply Syntax");
  var removeButton = buttonGroup.add("button", undefined, "Remove Syntax");

  applyButton.onClick = function () {
    processComps(
      compsList,
      true,
      textCheckbox.value,
      pathCheckbox.value,
      projectItemCheckbox.value
    );
  };

  removeButton.onClick = function () {
    processComps(
      compsList,
      false,
      textCheckbox.value,
      pathCheckbox.value,
      projectItemCheckbox.value
    );
  };

  dialog.show(); // This now works in a non-blocking mode since it's a palette
}

function processComps(
  compsList,
  apply,
  applyToText,
  applyToPath,
  applyToPrecomp
) {
  var project = app.project;

  // Begin the undo group
  app.beginUndoGroup(apply ? "Apply Syntax" : "Remove Syntax");

  function processLayer(layer) {
    // Apply or remove text syntax
    if (applyToText && layer instanceof TextLayer) {
      if (apply) {
        if (layer.name.indexOf("@") !== 0) {
          layer.name = "@" + layer.name;
        }
      } else {
        if (layer.name.indexOf("@") === 0) {
          layer.name = layer.name.substring(1);
        }
      }
    }

    // Apply or remove path syntax
    if (applyToPath && layer.source instanceof FootageItem) {
      if (apply) {
        if (layer.name.indexOf("$") !== 0) {
          layer.name = "$" + layer.name;
        }
      } else {
        if (layer.name.indexOf("$") === 0) {
          layer.name = layer.name.substring(1);
        }
      }
    }

    // Apply or remove precomp syntax
    if (applyToPrecomp && layer.source instanceof CompItem) {
      if (apply) {
        if (layer.name.indexOf("#") !== 0) {
          layer.name = "#" + layer.name;
        }
      } else {
        if (layer.name.indexOf("#") === 0) {
          layer.name = layer.name.substring(1);
        }
      }
    }
  }

  function processComp(comp) {
    for (var j = 1; j <= comp.numLayers; j++) {
      var layer = comp.layer(j);

      processLayer(layer); // Process the current layer

      // Recursively process layers inside precomps if any checkbox is selected
      if (
        layer.source instanceof CompItem &&
        (applyToText || applyToPath || applyToPrecomp)
      ) {
        processComp(layer.source); // Recursion to handle all nested precomps
      }
    }
  }

  // Process selected compositions
  for (var i = 0; i < compsList.items.length; i++) {
    var compName = compsList.items[i].text;
    if (compsList.items[i].selected) {
      for (var j = 1; j <= project.numItems; j++) {
        if (
          project.item(j) instanceof CompItem &&
          project.item(j).name === compName
        ) {
          processComp(project.item(j));
        }
      }
    }
  }

  // End the undo group
  app.endUndoGroup();
}

// Function definitions
function textPathNameExtract() {
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
    // var saveTXTButton = buttons.add("button", undefined, "Save as TXT", {
    //   name: "txt",
    // });

    // Define the cancel button behavior
    cancelButton.onClick = function () {
      dlg.close();
    };

    // Define the save button behavior for CSV
    saveCSVButton.onClick = function () {
      saveCompositions(listBox, "CSV");
      dlg.close();
    };

    // // Define the save button behavior for TXT
    // saveTXTButton.onClick = function () {
    //   saveCompositions(listBox, "TXT");
    //   dlg.close();
    // };

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
      if (!layer) continue; // Skip if layer is undefined

      var layerName = layer.name || ""; // Default to empty string if name is undefined

      // Recursively process precompositions first, regardless of the layer name
      if (layer.source instanceof CompItem) {
        var subCompData = {
          textLayers: {},
          fileLayers: {},
          specialLayers: {},
          hasGreaterThanLayer: false,
        };

        searchPrecomps(layer.source, subCompData, project);

        // Manually merge the results from the subcomposition search
        for (var key in subCompData.textLayers) {
          parentCompData.textLayers[key] = subCompData.textLayers[key];
        }
        for (var key in subCompData.fileLayers) {
          parentCompData.fileLayers[key] = subCompData.fileLayers[key];
        }
        for (var key in subCompData.specialLayers) {
          parentCompData.specialLayers[key] = subCompData.specialLayers[key];
        }
        hasGreaterThanLayer =
          hasGreaterThanLayer || subCompData.hasGreaterThanLayer;
      }

      // Skip layers not starting with '@', '#', or '$'
      if (
        layerName.length === 0 ||
        (layerName.indexOf("@") !== 0 &&
          layerName.indexOf("#") !== 0 &&
          layerName.indexOf("$") !== 0)
      ) {
        continue;
      }

      // Extract text from text layers only if they contain '@'
      if (layer instanceof TextLayer && layer.property("Source Text") != null) {
        if (layerName.indexOf("@") === 0) {
          var textSource = layer.property("Source Text").value;
          if (textSource) {
            var textContent = textSource.text.replace(/[\r\n]+/g, " ");
            parentCompData.textLayers[layerName] = textContent;
          }
        }
      }

      // Capture all layers that start with '$'
      if (layerName.indexOf("$") === 0) {
        hasGreaterThanLayer = true;
        var filePath = "no path retrieved";

        // Check if the layer's source has a file
        if (layer.source && layer.source.file) {
          filePath = layer.source.file.fsName;
        }

        parentCompData.fileLayers[layerName] = filePath;
      }

      // Check if layer name starts with '#'
      if (layerName.indexOf("#") === 0) {
        var sourceItemName = getSourceItemName(layer.source, project);
        parentCompData.specialLayers[layerName] =
          sourceItemName || "No source found";
      }
    }

    // Mark the composition with hasGreaterThanLayer flag
    parentCompData.hasGreaterThanLayer = hasGreaterThanLayer;
  }

  // Helper function to get the name of the source item from the project panel
  function getSourceItemName(layerSource, project) {
    if (!layerSource) {
      return null; // Return null if source is invalid
    }

    var sourceName = "No source found";
    for (var i = 1; i <= project.items.length; i++) {
      var item = project.items[i];
      if (item && item.id === layerSource.id) {
        sourceName = item.name;
        break;
      }
    }

    return sourceName; // Return the name of the project item or "No source found"
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

    for (var layerName in uniqueFileLayerNames) {
      if (uniqueFileLayerNames.hasOwnProperty(layerName)) {
        headers.push(layerName);
      }
    }

    for (var layerName in uniqueSpecialLayerNames) {
      if (uniqueSpecialLayerNames.hasOwnProperty(layerName)) {
        headers.push(layerName);
      }
    }

    var csvContent = headers.join(delimiter) + "\n";

    for (var i = 0; i < comps.length; i++) {
      var compObj = comps[i];
      if (!compObj) continue; // Skip if compObj is undefined
      var compRow = ['"' + (compObj.name || "").replace(/"/g, '""') + '"'];

      for (var j = 1; j < headers.length; j++) {
        var header = headers[j];
        var actualLayerName = header || ""; // Default to empty string if header is undefined
        var content = "";

        if (
          compObj.textLayersContent &&
          compObj.textLayersContent.hasOwnProperty(actualLayerName)
        ) {
          content = compObj.textLayersContent[actualLayerName];
        } else if (
          compObj.fileLayersContent &&
          compObj.fileLayersContent.hasOwnProperty(actualLayerName)
        ) {
          if (compObj.hasGreaterThanLayer) {
            content = compObj.fileLayersContent[actualLayerName];
          } else {
            content = "";
          }
        } else if (
          compObj.specialLayersContent &&
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
    var fileExtension = delimiter === "," ? "csv" : "txt";
    var file = new File(
      File.saveDialog("Save your file", "*." + fileExtension)
    );

    if (file) {
      // Ensure the file has the correct extension
      if (file.name.split(".").pop().toLowerCase() !== fileExtension) {
        file = new File(file.absoluteURI + "." + fileExtension);
      }

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
}

function compsFromSheet() {
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
    // Returns an object containing the suffix and a flag indicating whether to import suffix from CSV.
    function getPrecompSuffix() {
      var dialog = new Window("dialog", "Precomp Suffix Options");

      dialog.add(
        "statictext",
        undefined,
        "Choose how to provide the precomp suffix:"
      );

      var radioGroup = dialog.add("group");
      radioGroup.orientation = "column";

      var manualOption = radioGroup.add(
        "radiobutton",
        undefined,
        "Enter suffix manually (common precomps)"
      );
      var csvOption = radioGroup.add(
        "radiobutton",
        undefined,
        "Import from 'precomp suffix' column' (common precomps groups)"
      );

      manualOption.value = true; // Set default option

      // If manualOption is selected, show the input field
      var inputGroup = dialog.add("group");
      inputGroup.orientation = "row";

      inputGroup.add("statictext", undefined, "Suffix:");
      var input = inputGroup.add("edittext", undefined, "");
      input.characters = 20;

      // Disable input field when csvOption is selected
      manualOption.onClick = function () {
        input.enabled = true;
      };

      csvOption.onClick = function () {
        input.enabled = false;
      };

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
        var suffix = "";
        var useCsvSuffix = false;

        if (manualOption.value) {
          suffix = input.text.replace(/^\s+|\s+$/g, ""); // Trim whitespace
        } else {
          useCsvSuffix = true;
        }

        return { suffix: suffix, useCsvSuffix: useCsvSuffix };
      } else {
        return null; // Dialog cancelled
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

      // Adjust precompMap key to include suffix
      var compKey = comp.name + "_" + suffix;

      // Check if the composition has already been processed
      if (precompMap[compKey]) {
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

        // Handle precomps with "#" symbol without adding a suffix or duplicating them
        if (layerName.indexOf("#") === 0) {
          var columnName = layerName.substring(1);
          var precompName = rowData["#" + columnName]; // Precomp name from CSV
          var preComp = findProjectItemByName(precompName); // Find the precomp in the project by name

          if (preComp) {
            var originalStartTime = layer.startTime;
            var originalInPoint = layer.inPoint;
            var originalOutPoint = layer.outPoint;
            var originalDuration = layer.outPoint - layer.inPoint;
            var originalStretch = layer.stretch;
            var originalEnabled = layer.enabled;

            // Replace the layer with the existing precomp from the project panel
            layer.replaceSource(preComp, false);

            layer.startTime = originalStartTime;
            layer.inPoint = originalInPoint;
            layer.outPoint = originalOutPoint;
            layer.stretch = originalStretch;
            layer.enabled = originalEnabled;

            // Skip further processing for # layers, no suffix to be applied
            continue;
          } else {
            alert("Composition not found in project: " + precompName);
          }
        }

        // Handle precomps with symbols like @ and other cases where duplication is needed
        if (layer.source instanceof CompItem) {
          var preComp = layer.source;

          if (containsLayerWithSymbols(preComp, symbols)) {
            var preCompKey = preComp.name + "_" + suffix;

            if (!precompMap[preCompKey]) {
              var preCompDuplicate = preComp.duplicate();
              preCompDuplicate.name = preComp.name + suffix;
              preCompDuplicate.parentFolder = precompsFolder;

              // Mark the duplicated precomp in the precompMap
              precompMap[preCompKey] = preCompDuplicate;

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
            layer.replaceSource(precompMap[preCompKey], false);
            layer.enabled = true;
          }
        }
      }

      // Mark the top-level composition as processed to avoid reprocessing
      precompMap[compKey] = comp;
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

    // Function to get new composition names from the user or CSV
    function getCompNames(compsToRename, data) {
      var dialog = new Window("dialog", "Composition Naming Options");

      dialog.add(
        "statictext",
        undefined,
        "Choose how to name the new compositions:"
      );

      var radioGroup = dialog.add("group");
      radioGroup.orientation = "column";

      var manualOption = radioGroup.add(
        "radiobutton",
        undefined,
        "Enter names manually"
      );
      var csvOption = radioGroup.add(
        "radiobutton",
        undefined,
        "Use names from CSV (column 'new comp name')"
      );

      manualOption.value = true; // Set default option

      // Input field for manual entry
      var inputGroup = dialog.add("group");
      inputGroup.orientation = "column";

      inputGroup.add(
        "statictext",
        undefined,
        "Enter new names (one per line):"
      );
      var namesInput = inputGroup.add("edittext", undefined, "", {
        multiline: true,
      });
      namesInput.preferredSize = { width: 300, height: 200 };

      // Disable input field when csvOption is selected
      manualOption.onClick = function () {
        namesInput.enabled = true;
      };

      csvOption.onClick = function () {
        namesInput.enabled = false;
      };

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
        var newNames = [];
        var useCsvNames = false;

        if (manualOption.value) {
          newNames = namesInput.text.split("\n");
          for (var i = 0; i < newNames.length; i++) {
            newNames[i] = newNames[i].replace(/^\s+|\s+$/g, ""); // Trim whitespace
          }
        } else {
          useCsvNames = true;
          // Get names from CSV
          for (var i = 0; i < data.length; i++) {
            var newName = data[i]["new comp name"];
            newNames.push(newName || compsToRename[i].name); // Use existing name if not provided
          }
        }

        var renamedComps = [];
        for (var i = 0; i < compsToRename.length; i++) {
          renamedComps.push({
            oldName: compsToRename[i].name,
            newName: newNames[i] || compsToRename[i].name, // Retain the default name if not renamed
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
      var suffixData = getPrecompSuffix();

      if (!suffixData) {
        // User cancelled the dialog
        return;
      }

      // Store the suffix and whether to use CSV suffix
      var suffix = suffixData.suffix;
      var useCsvSuffix = suffixData.useCsvSuffix;

      var outputFolderName = "00_Generated Comps";
      var precompsFolderName = "Precomps";
      var importFolderName = "Imported Files";

      var data = parseDocument(filePath);
      var messages = [];
      var existingCompNames = [];
      var precompMap = {}; // Keep track of precomps
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

            var rowSuffix = suffix;
            if (useCsvSuffix) {
              rowSuffix = data[i]["precomp suffix"] || ""; // Get suffix from CSV or default to empty string
            }

            var newComp = createCompFromData(
              data[i],
              mainComp,
              outputFolder,
              precompsFolder,
              importFolder,
              existingCompNames,
              precompMap,
              rowSuffix
            );
            compsToRename.push(newComp);
          } else {
            var message = "Composition not found: " + compName;
            alert(message);
            messages.push(message);
          }
        }
      }

      // Get new composition names from the user or CSV
      var renamedComps = getCompNames(compsToRename, data);

      if (renamedComps) {
        for (var i = 0; i < renamedComps.length; i++) {
          var compInfo = renamedComps[i];
          var comp = findProjectItemByName(compInfo.oldName);
          if (comp) {
            comp.name = compInfo.newName;
          }
        }
      } else {
        // User canceled the renaming process
        return;
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

// Example usage
var myToolsPanel = createUI(this);

if (myToolsPanel instanceof Window) {
  myToolsPanel.center();
  myToolsPanel.show();
} else {
  myToolsPanel.layout.layout(true);
}
