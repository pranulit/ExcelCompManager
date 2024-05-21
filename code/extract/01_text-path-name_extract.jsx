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
  var listBox = dlg.add("listbox", [0, 0, 380, 300], [], { multiselect: true });
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
    if (compItem.name.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1) {
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

  for (var k = 1; k <= comp.numLayers; k++) {
    var layer = comp.layer(k);

    // Recursively handle precompositions
    if (layer.source instanceof CompItem) {
      searchPrecomps(layer.source, parentCompData, project);
    }

    // Extract text from text layers
    if (layer instanceof TextLayer && layer.property("Source Text") != null) {
      var textSource = layer.property("Source Text").value;
      if (textSource) {
        var textContent = textSource.text.replace(/[\r\n]+/g, " ");
        parentCompData.textLayers[layer.name] = textContent;
      }
    }

    // Capture all layers that start with ">", even if they don't have a file
    if (layer.name.substring(0, 1) === ">") {
      // If the layer has a source file, store the path; otherwise, note the absence of a path
      var filePath =
        layer.source && layer.source.file
          ? layer.source.file.fsName
          : "no path retrieved";
      parentCompData.fileLayers[layer.name] = filePath;
    }

    // Check if layer name starts with "#"
    if (layer.name.charAt(0) === "#") {
      // Extract and store the source name if the layer has a source
      if (layer.source) {
        var sourceName = getSourceName(layer.source.name, project);
        if (sourceName) {
          parentCompData.specialLayers[layer.name] = sourceName;
        } else {
          // Handle cases where source name is not found
          parentCompData.specialLayers[layer.name] = "Source not found";
        }
      } else {
        // Handle cases where there is no source
        parentCompData.specialLayers[layer.name] = "No source associated";
      }
    }
  }
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

    var parentCompData = { textLayers: {}, fileLayers: {}, specialLayers: {} };
    searchPrecomps(compItem, parentCompData, project);

    var compObj = {
      name: compItem.name,
      textLayersContent: parentCompData.textLayers,
      fileLayersContent: parentCompData.fileLayers,
      specialLayersContent: parentCompData.specialLayers,
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
      var isFilePath = header.indexOf(" File Path") > -1;
      var actualLayerName = isFilePath
        ? header.replace(" File Path", "")
        : header;
      var content = "";

      if (compObj.textLayersContent.hasOwnProperty(actualLayerName)) {
        content = compObj.textLayersContent[actualLayerName];
      } else if (compObj.fileLayersContent.hasOwnProperty(actualLayerName)) {
        content = compObj.fileLayersContent[actualLayerName];
      } else if (compObj.specialLayersContent.hasOwnProperty(actualLayerName)) {
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
