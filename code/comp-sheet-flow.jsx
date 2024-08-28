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
  var dialog = new Window("dialog", "Apply / Remove Syntax");
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
  var projectItemCheckbox = optionsGroup.add(
    "checkbox",
    undefined,
    "# (precomps by default)"
  );

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
    // Refresh the UI or give feedback without closing the dialog
  };

  removeButton.onClick = function () {
    processComps(
      compsList,
      false,
      textCheckbox.value,
      pathCheckbox.value,
      projectItemCheckbox.value
    );
    // Refresh the UI or give feedback without closing the dialog
  };

  dialog.show();
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

  function processComp(comp) {
    for (var j = 1; j <= comp.numLayers; j++) {
      var layer = comp.layer(j);

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

      // Apply or remove precomp syntax and recurse
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
        // Recursively process precompositions
        processComp(layer.source);
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
  // Implementation goes here
}

function compsFromSheet() {
  // Implementation goes here
}

// Example usage
var myToolsPanel = createUI(this);

if (myToolsPanel instanceof Window) {
  myToolsPanel.center();
  myToolsPanel.show();
} else {
  myToolsPanel.layout.layout(true);
}
