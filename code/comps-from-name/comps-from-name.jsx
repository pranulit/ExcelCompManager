(function() {
    var myPanel = (this instanceof Panel) ? this : new Window("palette", "Create Compositions", undefined, {resizeable:true});

    if (myPanel != null) {
        var textGroup = myPanel.add("group");
        textGroup.orientation = "row";
        textGroup.add("statictext", undefined, "Enter Composition Names:");

        var inputGroup = myPanel.add("group");
        inputGroup.orientation = "row";
        var compsInput = inputGroup.add("edittext", undefined, "", {multiline: true});
        compsInput.size = [300, 100];

        var fpsGroup = myPanel.add("group");
        fpsGroup.orientation = "row";
        fpsGroup.add("statictext", undefined, "Select FPS:");
        var fpsDropdown = fpsGroup.add("dropdownlist", undefined, ["24", "25", "30", "60"]);
        fpsDropdown.selection = 1;

        var buttonGroup = myPanel.add("group");
        buttonGroup.orientation = "row";
        var createBtn = buttonGroup.add("button", undefined, "Create Compositions");
        createBtn.onClick = function() {
            createCompositions(compsInput.text, parseFloat(fpsDropdown.selection.text));
        }
    }

    function createCompositions(compsInput, fps) {
        var lines = compsInput.split("\n");
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
                if (line !== "") {
                    var parts = line.split("_");
                    var duration = parts[2].match(/(\d+)s/);
                    var formatMatch = parts[2].match(/(\d+x\d+)/);
                    var format = formatMatch ? formatMatch[1] : defaultFormat;

                    if (duration && compSettings[format]) {
                        var comp = app.project.items.addComp(parts[0], compSettings[format].width, compSettings[format].height, 1, parseFloat(duration[1]), fps);
                        comp.parentFolder = mainFolder; // Assign composition to the main folder
                        comp.name = line;

                        var solid = comp.layers.addSolid([1, 1, 1], ">offline", comp.width, comp.height, 1);
                        solid.name = ">offline";
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
