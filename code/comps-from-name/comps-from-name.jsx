(function() {
    var myPanel = (this instanceof Panel) ? this : new Window("palette", "Create Compositions", undefined, {resizeable: true});

    if (myPanel != null) {
        var mainGroup = myPanel.add("group");
        mainGroup.orientation = "column";
        mainGroup.alignChildren = "fill";

        // Group for checkbox
        var checkboxGroup = mainGroup.add("group");
        checkboxGroup.orientation = "row";
        var addOfflineCheckbox = checkboxGroup.add("checkbox", undefined, "Add Offline File");

        // Group for the labels and input boxes
        var inputGroup = mainGroup.add("group");
        inputGroup.orientation = "row";
        inputGroup.alignChildren = "top";
        inputGroup.alignment = ["center", "top"];

        // Subgroup for composition names
        var compNameGroup = inputGroup.add("group");
        compNameGroup.orientation = "column";
        compNameGroup.alignChildren = "center";
        compNameGroup.add("statictext", undefined, "Add Composition Names:");
        var compsInput = compNameGroup.add("edittext", undefined, "", {multiline: true});
        compsInput.size = [300, 100];

        // Subgroup for offline file input
        var offlineFileGroup = inputGroup.add("group");
        offlineFileGroup.orientation = "column";
        offlineFileGroup.alignChildren = "center";
        offlineFileGroup.add("statictext", undefined, "Add Offline File:");
        var offlineFileInput = offlineFileGroup.add("edittext", undefined, "", {multiline: true});
        offlineFileInput.size = [300, 100];
        offlineFileGroup.visible = false;

        addOfflineCheckbox.onClick = function() {
            var visible = addOfflineCheckbox.value;
            offlineFileGroup.visible = visible;
            compNameGroup.alignment = visible ? ["left", "top"] : ["center", "top"];
            inputGroup.layout.layout(true); // Adjust the layout when visibility changes
        };

        // Group for FPS selection and button
        var bottomGroup = mainGroup.add("group");
        bottomGroup.orientation = "column";
        bottomGroup.alignChildren = "fill";

        var fpsGroup = bottomGroup.add("group");
        fpsGroup.orientation = "row";
        fpsGroup.add("statictext", undefined, "Select FPS:");
        var fpsDropdown = fpsGroup.add("dropdownlist", undefined, ["24", "25", "30", "60"]);
        fpsDropdown.selection = 1;

        var buttonGroup = bottomGroup.add("group");
        buttonGroup.orientation = "row";
        var createBtn = buttonGroup.add("button", undefined, "Create Compositions");
        createBtn.onClick = function() {
            var offlineText = offlineFileInput.visible ? offlineFileInput.text : "";
            alert("compsInput.text: " + compsInput.text);
            alert("offlineFileInput.text: " + offlineText);
            createCompositions(compsInput.text, offlineText, parseFloat(fpsDropdown.selection.text));
        };
    }

    function createCompositions(compsInput, offlineInputText, fps) {
        alert("createCompositions called");
        alert("compsInput: " + compsInput);
        alert("offlineInputText: " + offlineInputText);
        alert("fps: " + fps);

        var lines = compsInput.split("\n");
        alert("lines: " + lines);

        var mainFolder = app.project.items.addFolder("main"); // Create the main folder

        var compSettings = {
            "16x9": {width: 1920, height: 1080},
            "1x1": {width: 1080, height: 1080},
            "4x5": {width: 1080, height: 1350},
            "9x16": {width: 1080, height: 1920}
        };
        var defaultFormat = "16x9";

        app.beginUndoGroup("Create Compositions");

        try {
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i].replace(/^\s+|\s+$/g, ''); // Trim whitespace using regex
                alert("Processing line: " + line);
                if (line !== "") {
                    var parts = line.split("_");
                    alert("parts: " + parts);
                    if (parts.length < 3) {
                        alert("Invalid format for line: " + line);
                        continue;
                    }
                    var duration = parts[2].match(/(\d+)s/);
                    var formatMatch = parts[2].match(/(\d+x\d+)/);
                    var format = formatMatch ? formatMatch[1] : defaultFormat;
                    alert("duration: " + duration);
                    alert("format: " + format);

                    if (duration && compSettings[format]) {
                        var comp = app.project.items.addComp(parts[0], compSettings[format].width, compSettings[format].height, 1, parseFloat(duration[1]), fps);
                        comp.parentFolder = mainFolder; // Assign composition to the main folder
                        comp.name = line;

                        var solidName = offlineInputText.trim() !== "" ? offlineInputText : ">offline";
                        alert("solidName: " + solidName);
                        var solid = comp.layers.addSolid([1, 1, 1], solidName, comp.width, comp.height, 1);
                        solid.name = solidName;
                    } else {
                        alert("Could not determine valid duration/format for: " + line);
                    }
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
