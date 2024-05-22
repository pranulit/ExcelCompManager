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
    var createBtn = buttonGroup.add("button", undefined, "Create Compositions");
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
