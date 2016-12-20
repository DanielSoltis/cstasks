#target Illustrator

/********
CSTasks is a set of basic functions for selecting, moving, copying and exporting objects in Illustrator (might be generalizable to other CS programs?)
It uses an array [x,y] to indicate the top left corner of an object, collection, artboard, etc., to provide a consistent and
easy-to-understand frame of reference for moving objects.

Tested using Creative Cloud 2017

Note that there's some wonkiness around grouping and ungrouping. In general, this scripting gets strange quickly.
********/

var CSTasks = (function(){
 
    var tasks = {};
   
    /********************
    POSITION AND MOVEMENT
    ********************/

    //takes an artboard
    //returns its left top corner as an array [x,y]
    tasks.getArtboardCorner = function(artboard){ 
        var corner = [artboard.artboardRect[0], artboard.artboardRect[1]];
        return corner;
    };

    //takes an artboards array
    //returns the leftmost and topmost position of all the artboards as an array [x,y]
    tasks.getArtboardsTopLeft = function(artboards){
        var pos = [Infinity, -Infinity]; //numbering goes up left to right, and down top to bottom
        for (var i = 0; i < artboards.length; i++){
            var rect = artboards[i].artboardRect;
            if (rect[0] < pos[0]) pos[0] = rect[0];
            if (rect[1] > pos[1]) pos[1] = rect[1];
        }
        return pos;
    };

    //takes the document, a collection of objects (e.g. selection, PathItems, GroupItems) and a destination array [x,y]
    //returns the leftmost and topmost position of all the items as an array [x,y]
    tasks.getCollectionTopLeft = function(doc, collection){
        var tempGroup = tasks.createGroup(doc, collection);
        var pos = tempGroup.position;
        tasks.ungroupOnce(tempGroup);
        return pos;
    };

    //takes an array [x,y] for an item's position and an array [x,y] for the position of a reference point
    //returns an aray [x,y] for the offset between the two points
    tasks.getOffset = function(itemPos, referencePos){
        var offset = [itemPos[0] - referencePos[0], itemPos[1] - referencePos[1]];
        return offset;
    };

    //takes the document, a collection of objects (e.g. selection, PathItems, GroupItems) and a destination array [x,y]
    //moves the collection to the specified destination
    tasks.translateCollectionTo = function(doc, collection, destination){
        var tempGroup = tasks.createGroup(doc, collection);
        tasks.translateObjectTo(tempGroup, destination);
        tasks.ungroupOnce(tempGroup);
    };

    //takes an object (e.g. group) and a destination array [x,y]
    //moves the group to the specified destination
    tasks.translateObjectTo = function(object, destination){
        var offset = tasks.getOffset(object.position, destination);
        object.translate(-offset[0],-offset[1]);
    };
    
    //takes a document and index of an artboard
    //deletes everything on that artboard
    tasks.clearArtboard = function(doc, index){ //clears an artboard at the given index
		doc.selection = null;
		doc.artboards.setActiveArtboardIndex(index);
		doc.selectObjectsOnActiveArtboard();
		var sel = doc.selection; // get selection 
		for (k=0; k<sel.length; k++) {  
			sel[k].remove(); 
		}
	};


    /*********************************
    SELECTING, GROUPING AND UNGROUPING
    **********************************/
    
    //take a document
    //returns a selection with everything in that document
    tasks.selectEverything = function(doc){
        doc.selection = null;
        for ( i = 0; i < doc.pathItems.length; i++ ) {
            doc.pathItems[i].selected = true;
        }
        return doc.selection;
    };    
    
    //takes a document and the index of an artboard in that document's artboards array
    //returns a selection of all the objects on that artboard
    tasks.selectContentsOnArtboard = function(doc, i){
        doc.selection = null;	
        doc.artboards.setActiveArtboardIndex(i);
        doc.selectObjectsOnActiveArtboard();
        return doc.selection; 
    };

    //takes a document and a collection of objects (e.g. selection)
    //returns a group made from that collection
    tasks.createGroup = function(doc, collection){
        var newGroup = doc.groupItems.add();
        for (k=0; k<collection.length; k++) {  
            collection[k].moveToBeginning(newGroup);
        } 
        return newGroup;
    };

    //takes a group
    //ungroups that group at the top layer (no recursion for nested groups)
    tasks.ungroupOnce = function(group){
        for (i=group.pageItems.length-1; i>=0; i--)  {
            group.pageItems[i].move(group.pageItems[i].layer, ElementPlacement.PLACEATEND);  
        }
    };

    /****************************
    CREATING AND SAVING DOCUMENTS
    *****************************/

    //take a source document and a colorspace (e.g. DocumentColorSpace.RGB)
    //opens and returns a new document with the source document's units and the specified colorspace
    tasks.newDocument = function(sourceDoc, colorSpace){
        var preset = new DocumentPreset();
        preset.colorMode = colorSpace;
        preset.units = sourceDoc.rulerUnits;
        
        var newDoc = app.documents.addDocument(colorSpace, preset);
        newDoc.pageOrigin = sourceDoc.pageOrigin;
        newDoc.rulerOrigin = sourceDoc.rulerOrigin;

        return newDoc;
    };
    
    //take a source document, artboard index, and a colorspace (e.g. DocumentColorSpace.RGB)
    //opens and returns a new document with the source document's units and specified artboard, the specified colorspace 
    tasks.duplicateArtboardInNewDoc = function(sourceDoc, artboardIndex, colorspace){
        var rectToCopy = sourceDoc.artboards[artboardIndex].artboardRect;
        var newDoc = tasks.newDocument(sourceDoc, colorspace);
        newDoc.artboards.add(rectToCopy);	
        newDoc.artboards.remove(0);
        return newDoc;
    };

    //take a source document and a colorspace (e.g. DocumentColorSpace.RGB)
    //opens and returns a new document with the source document's units and artboards, the specified colorspace 
    tasks.duplicateArtboardsInNewDoc = function(sourceDoc, colorspace){
        var rectsToCopy = new Array(sourceDoc.artboards.length);
        for (var i = 0; i < sourceDoc.artboards.length; i++){
            rectsToCopy[i] = sourceDoc.artboards[i].artboardRect;
        }

        var newDoc = tasks.newDocument(sourceDoc, colorspace);
        
        for (var i = 0; i < rectsToCopy.length; i++){
            var thisRect = rectsToCopy[i];
            newDoc.artboards.add(thisRect);
        }		
        newDoc.artboards.remove(0);
        return newDoc;
    };
    
    //takes a selection and a destination document
    //duplicates the selection into a new document (arbitrary position)
    tasks.copySelectionToNewDoc = function(selection, destDoc){
		var newItem;
		if (selection.length > 0 ) {
			for (var i = 0; i < selection.length; i++ ) {
				selection[i].selected = false;
				newItem = selection[i].duplicate(destDoc, ElementPlacement.PLACEATEND);
			}
		}
		else {
			selection.selected = false;
			newItem = selection.parent.duplicate(destDoc, ElementPlacement.PLACEATEND);
		}
	};


    //take a document and desired colorspace (e.g. DocumentColorSpace.RGB)
    //creates and returns a new document in that colorspace with all artboards and contents duplicated
    tasks.duplicateDocument = function(doc, colorspace){
        var sel = tasks.selectEverything(doc);
        var selGroup = tasks.createGroup(doc,sel);
        var abPos = tasks.getArtboardsTopLeft(doc.artboards);
        var offset = tasks.getOffset(selGroup.position,abPos);
        var newDoc = tasks.duplicateArtboardsInNewDoc(doc, colorspace);
        var newGroup = selGroup.duplicate(newDoc.layers[selGroup.layer.name],ElementPlacement.PLACEATEND);
        var newAbPos = tasks.getArtboardsTopLeft(newDoc.artboards);
        tasks.translateObjectTo(newGroup,[newAbPos[0] + offset[0], newAbPos[1] + offset[1]]);
        tasks.ungroupOnce(selGroup);
        tasks.ungroupOnce(newGroup);
        return newDoc;
    };
    
    //take a document, export file, initial width and desired width
    //exports a PNG scaled proportionally to the desired width
    tasks.scaleAndExportPNG = function(doc, destFile, startWidth, desiredWidth) {
		var scaling = 100.0*desiredWidth/startWidth;
		var options = new ExportOptionsPNG24();
		options.antiAliasing = true;
		options.transparency = true; 
		options.artBoardClipping = true;
		options.horizontalScale = scaling;
		options.verticalScale = scaling;	

		doc.exportFile(destFile, ExportType.PNG24, options);
	};

    //takes left x, top y, width, and height
    //returns a Rect that can be used to create an artboard
    tasks.newRect = function(x, y, width, height) {
        var rect = [];
        rect[0] = x;
        rect[1] = -y;
        rect[2] = width + x;
        rect[3] = -(height + y);
        return rect;
    };

	/***
	TEXT
	****/
	
	//takes a text frame and a string with the desired font name
	//sets the text frame to the desired font or alerts if not found
	tasks.setFont = function(textRef, desiredFont){
		var foundFont = false;
		for (var i = 0; i < textFonts.length; i++){
			if (textFonts[i].name == desiredFont) {
				textRef.textRange.characterAttributes.textFont = textFonts[i];
				foundFont = true;
				break;
			}		
		}
		if (!foundFont) alert("Didn't find the font. Please check if the font is installed or check the script to make sure the font name is right.");
	};
	
	//takes a document, message string, position array and font size
    //creates a text frame with the message
    tasks.createTextFrame = function(doc, message, pos, size){
        var textRef = doc.textFrames.add();
        textRef.contents = message;
        textRef.left = pos[0];
        textRef.top = pos[1];
        textRef.textRange.characterAttributes.size = size;
    };

    /*********
    RGB Colors
    **********/
    //for future, may be better to combine these two sets of functions and use color space as an argument
    
    //compares two RGB colors
    //returns true if they match, false if they do not
    tasks.matchColorsRGB = function(color1, color2){ //compares two colors to see if they match
		if (Math.abs(color1.red - color2.red) < 1 &&
		Math.abs(color1.green - color2.green) < 1 &&
		Math.abs(color1.blue - color2.blue) < 1) { //can't do equality because it adds very small decimals
			return true;
		}
        return false;
    };
    
    //take a single RGBColor and an array of corresponding RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns the index in the array if it finds a match, otherwise returns -1
    tasks.matchToPaletteRGB = function(color, matchArray){ //compares a single color RGB color against RGB colors in [[RGB],[CMYK]] array
        for (var i = 0; i < matchArray.length; i++){
            if (Math.abs(color.red - matchArray[i][0].red) < 1 &&
            Math.abs(color.green - matchArray[i][0].green) < 1 &&
            Math.abs(color.blue - matchArray[i][0].blue) < 1) { //can't do equality because it adds very small decimals
                return i;
            }
        }
        return -1;
    };
      
    //takes a collection of pathItems and an array of specified RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns an array with the values corresponding to the indices of the relevant color in the palette (or -1 if no match)
    tasks.indexPaletteRGB = function(pathItems, matchArray){
        var colorIndex = new Array(pathItems.length);
        for ( i = 0; i < pathItems.length; i++ ) {
            var itemColor = pathItems[i].fillColor;
            colorIndex[i] = tasks.matchToPaletteRGB(itemColor, matchArray);
        }
        return colorIndex;
    };
    
    //takes a pathItems array, startColor and endColor
    //converts all pathItems with startColor into endColor
    tasks.convertColorRGB = function(pathItems, startColor, endColor){
        for ( i = 0; i < pathItems.length; i++ ) {
            if (tasks.matchColorsRGB(pathItems[i].fillColor, startColor)) pathItems[i].fillColor = endColor;
        }
    };
    
 
    /**********
    CMYK Colors
    ***********/
    
    //compares two CMYK colors
    //returns true if they match, false if they do not
    tasks.matchColorsCMYK = function(color1, color2){ //compares two colors to see if they match
		if (Math.abs(color1.cyan - color2.cyan) < 1 &&
		Math.abs(color1.magenta - color2.magenta) < 1 &&
		Math.abs(color1.yellow - color2.yellow) < 1 &&
		Math.abs(color1.black - color2.black) < 1) { //can't do equality because it adds very small decimals
			return true;
		}
        return false;
    };
    
    //take a single CMYKColor and an array of corresponding RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns the index in the array if it finds a match, otherwise returns -1
    tasks.matchToPaletteCMYK = function(color, matchArray){ //compares a single color RGB color against RGB colors in [[RGB],[CMYK]] array
        for (var i = 0; i < matchArray.length; i++){
            if (Math.abs(color.red - matchArray[i][1].red) < 1 &&
            Math.abs(color.green - matchArray[i][1].green) < 1 &&
            Math.abs(color.blue - matchArray[i][1].blue) < 1) { //can't do equality because it adds very small decimals
                return i;
            }
        }
        return -1;
    };
    
    //takes a collection of pathItems and an array of specified RGB and CMYK colors [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...]
    //returns an array with an index to the CMYK color if it is in the array
    tasks.indexPaletteCMYK = function(pathItems, matchArray){
        var colorIndex = new Array(pathItems.length);
        for ( i = 0; i < pathItems.length; i++ ) {
            var itemColor = pathItems[i].fillColor;
            colorIndex[i] = tasks.matchToPaletteCMYK(itemColor, matchArray);
        }
        return colorIndex;
    };
    
    //takes a pathItems array, startColor and endColor and converts all pathItems with startColor into endColor
    tasks.convertColorCMYK = function(pathItems, startColor, endColor){
        for ( i = 0; i < pathItems.length; i++ ) {
            if (tasks.matchColorsCMYK(pathItems[i].fillColor, startColor)) pathItems[i].fillColor = endColor;
        }
    };
    
    /*******************************************************
    Multiple color spaces or conversion between color spaces
    *******************************************************/
    
    //takes two equal-length arrays of corresponding colors [[R,G,B], [R2,G2,B2],...] and [[C,M,Y,K],[C2,M2,Y2,K2],...] (fairly human readable)
    //returns an array of ColorElements [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...] (usable by the script for fill colors etc.)
    tasks.initializeColorPalette = function(RGBArray, CMYKArray){
        var colors = new Array(RGBArray.length);
        for (var i = 0; i < RGBArray.length; i++){
            var rgb = new RGBColor();
            rgb.red = RGBArray[i][0];
            rgb.green = RGBArray[i][1];
            rgb.blue = RGBArray[i][2];
        
            var cmyk = new CMYKColor();
            cmyk.cyan = CMYKArray[i][0];
            cmyk.magenta = CMYKArray[i][1];	
            cmyk.yellow = CMYKArray[i][2];	
            cmyk.black = CMYKArray[i][3];		
            
            colors[i] = [rgb,cmyk];	
        }
        return colors;
    };
    
    //takes a pathItems array, endColor and opacity and converts all pathItems into endColor at the specified opacity
    tasks.changeAllToColor = function(pathItems, endColor, opcty){
         for ( i = 0; i < pathItems.length; i++ ) {
            pathItems[i].fillColor = endColor;
            pathItems[i].opacity = opcty;
        }
    };
 
    //takes a doc, collection of pathItems, an array of specified colors and an array of colorIndices
    //converts the fill colors to the indexed CMYK colors and adds a text box with the unmatched colors
    //Note that this only makes sense if you've previously indexed the same path items and haven't shifted their positions in the pathItems array
    //This lets you index the colors while they are in a document with RGB color space and do the conversion in a new document in CMYK color space
    tasks.convertToCMYKPalette = function(doc, pathItems, colorArray, colorIndex){
        var unmatchedColors = [];
        for ( i = 0; i < pathItems.length; i++ ) {
            if (colorIndex[i] >=0 && colorIndex[i] < colorArray.length) pathItems[i].fillColor = colorArray[colorIndex[i]][1];
            else {
                var unmatchedColor = "(" + pathItems[i].fillColor.red + ", " + pathItems[i].fillColor.green + ", " + pathItems[i].fillColor.blue + ")";
                unmatchedColors.push(unmatchedColor);	
            }
        }
        if (unmatchedColors.length > 0){  
        	alert("One or more colors don't match the brand palette and weren't converted.");        
            unmatchedColors = tasks.unique(unmatchedColors);
            var unmatchedString = "Unconverted colors:";
            for (var i = 0; i < unmatchedColors.length; i++){
                unmatchedString = unmatchedString + "\n" + unmatchedColors[i];
            }
            var errorMsgPos = [Infinity, Infinity]; //gets the bottom left of all the artboards
            for (var i = 0; i < doc.artboards.length; i++){
            var rect = doc.artboards[i].artboardRect;
				if (rect[0] < errorMsgPos[0]) errorMsgPos[0] = rect[0];
				if (rect[3] < errorMsgPos[1]) errorMsgPos[1] = rect[3];
        	}
        	errorMsgPos[1] = errorMsgPos[1] - 20;
        	
            tasks.createTextFrame(doc, unmatchedString, errorMsgPos, 18);
        }
    };
    
    //takes a pathItems array, startColor and endColor and converts all pathItems with startColor into endColor
    tasks.convertColorCMYK = function(pathItems, startColor, endColor){
        for ( i = 0; i < pathItems.length; i++ ) {
            if (tasks.matchColorsCMYK(pathItems[i].fillColor, startColor)) pathItems[i].fillColor = endColor;
        }
    };

    //takes an array
    //returns a sorted array with only unique elements (not strictly color conversion but used in the previous function)
    tasks.unique = function(a) {
        if (a.length > 0){
            sorted =  a.sort();
            uniq = [sorted[0]];
            for (var i = 1; i < sorted.length; i++){
                if (sorted[i] != sorted[i-1]) uniq.push(sorted[i]);
            }
            return uniq;
        }
        return [];
    };

    return tasks;
}());
