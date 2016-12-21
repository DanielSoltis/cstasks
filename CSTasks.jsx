//#target Illustrator
//UNCOMMENT THE LINE ABOVE

/********
CSTasks is a set of functions for selecting, moving, copying and exporting objects in Illustrator (might be generalizable to other CS programs?)
It uses an array [x,y] to indicate the top left corner of an object, collection, artboard, etc., to provide a consistent and
easy-to-understand frame of reference for moving objects.

Tested using Creative Cloud 2017

Note that there's some wonkiness around grouping and ungrouping. In general, this scripting gets strange quickly if you string together too many tasks.
********/

var CSTasks = (function(){
 
    var tasks = {};    
    
    /*********************
    SELECTING AND GROUPING
    **********************/
      
    /**
     * @function selectEverything
	 * @description Selects everything in a document that is unlocked and returns an array of all the selected objects
	 * @param {Document} doc - The document (should be open and active)
	 * @return {array} selection - Array of all the selected objects
	 */
    tasks.selectEverything = function(doc){
        doc.selection = null;
        for (var i = 0; i < doc.pathItems.length; i++ ) {
            if (!doc.pathItems[i].layer.locked && !doc.pathItems[i].locked) doc.pathItems[i].selected = true;
        }
        return doc.selection;
    };    
    
    /**
     * @function selectContentsOnArtboard
	 * @description Selects everything on a specified artboard and returns an array of all the selected objects
	 * @param {Document} doc - The document (should be open and active)
	 * @param {number} i - Index of the artboard
	 * @return {array} selection - Array of all the selected objects
	 */
    tasks.selectContentsOnArtboard = function(doc, i){
        doc.selection = null;
        if (i >= 0 && i < doc.artboards.length){	
        	doc.artboards.setActiveArtboardIndex(i);
        	doc.selectObjectsOnActiveArtboard();
        }
        else alert("There is no artboard with the index " + i);
        
        return doc.selection; 
    };

    /**
     * @function groupSelection
	 * @description Creates and returns a group from an array of selected objects
	 * @param {Document} doc - The document (should be open and active)
	 * @param {Array} sel - Array of selected objects
	 * @return {Group} group - The grouped selection
	 */

    tasks.groupSelection = function(doc, sel){
        var newGroup = doc.groupItems.add();
        for (var i = 0; i<sel.length; i++) {  
            sel[i].moveToBeginning(newGroup);
        } 
        return newGroup;
    };

    /**
     * @function ungroupOnce
	 * @description Ungroups a group (no recursion for nested groups)
	 * @param {Group} group
	 */
    tasks.ungroupOnce = function(group){
        for (var i=group.pageItems.length-1; i>=0; i--)  {
            group.pageItems[i].move(group.pageItems[i].layer, ElementPlacement.PLACEATEND);  
        }
    };
    
    /**
     * @function clearArtboard
	 * @description Deletes every unlocked object on the specified artboard
	 * @param {Document} doc - The  document (should be open and active)
	 * @param {number} i - Index of the artboard
	 */ 
    tasks.clearArtboard = function(doc, index){ //clears an artboard at the given index
		doc.selection = null;
		doc.artboards.setActiveArtboardIndex(index);
		doc.selectObjectsOnActiveArtboard();
		var sel = doc.selection; // get selection 
		for (var i = 0; i<sel.length; i++) {  
			sel[i].remove(); 
		}
	};
    
    /********************
    POSITION AND MOVEMENT
    ********************/

	/**
     * @function getArtboardTopLeft
	 * @description Takes an artboard and returns the position of its top left corner
	 * @param {artboard} artboard - The Artboard object
	 * @return {array} position - The position of the top left corner of the artboard as an [x,y] array
	 */
    tasks.getArtboardTopLeft = function(artboard){ 
        var pos = [artboard.artboardRect[0], artboard.artboardRect[1]];
        return pos;
    };

	/**
     * @function getArtboardsTopLeft
	 * @description Takes an array of artboards and returns the position of their leftmost and topmost corner
	 * @param {Artboards} artboards - The Artboards array 
	 * @return {array} position - The position of the topmost leftmost corner of all the artboards as an [x,y] array
	 */
    tasks.getArtboardsTopLeft = function(artboards){
        var pos = [Infinity, -Infinity]; //numbering goes up left to right, and down top to bottom
        for (var i = 0; i < artboards.length; i++){
            var rect = artboards[i].artboardRect;
            if (rect[0] < pos[0]) pos[0] = rect[0];
            if (rect[1] > pos[1]) pos[1] = rect[1];
        }
        return pos;
    };

	/**
     * @function getCollectionTopLeft
	 * @description Takes an array of objects (e.g. selection, PathItems, GroupItems) 
	 * @param {array} collection - The array of objects
	 * @return {array} position - The position of the topmost leftmost corner of all the objects as an [x,y] array
	 */
    tasks.getCollectionTopLeft = function(collection){
    	var pos = [Infinity, -Infinity]; //numbering goes up left to right, and down top to bottom
        for (var i = 0; i < collection.length; i++){
        	if (collection[i].position[0] < pos[0]) pos[0] = collection[i].position[0];
        	if (collection[i].position[1] > pos[1]) pos[1] = collection[i].position[1];
        }
        return pos;
    };

	/**
     * @function getOffset
	 * @description Returns the offset between two positions as an [x,y] array.
	 * Useful for preserving position when duplicating objects into new documents
	 * @param {array} pos1 - A position as an [x,y] array
	 * @param {array} pos2 - Another position as an [x,y] array
	 * @return {array} offset - The offset between the two positions as an [x,y] array
	 */
    tasks.getOffset = function(pos1, pos2){
        var offset = [pos1[0] - pos2[0], pos1[1] - pos2[1]];
        return offset;
    };

	/**
     * @function translateSelectionTo
	 * @description Takes an array of selected objects and translates their position so the top left corner
	 * is at the specified destination. Preserves relative positions of objects.
	 * @param {array} sel - The array of selected objects
	 * @param {array} destination - The destination position as an [x,y] array
	 */    
    tasks.translateSelectionTo = function(sel, destination){ 
    	var pos = tasks.getCollectionTopLeft(sel);
    	var offset = tasks.getOffset(destination, pos);
        for (var i = 0; i<sel.length; i++) {  
        	sel[i].translate(offset[0],offset[1]);
        } 
    };

	/**
     * @function translateObjectTo
	 * @description Takes an object and translates its position so the top left corner is at the specified destination
	 * @param sel - The object (such as a pathItem or group)
	 * @param {array} destination - The destination position as an [x,y] array
	 */ 
    tasks.translateObjectTo = function(object, destination){
        var offset = tasks.getOffset(destination, object.position);
        object.translate(offset[0],offset[1]);
    };

    /******************************************
    CREATING DOCUMENTS AND DUPLICATING CONTENTS
    ******************************************/

    /**
     * @function newDocument
	 * @description Tales a source document and colorspace and opens and returns new document with the same units in the specified colorspace.
	 * @param {Document} sourceDoc - The source document
	 * @param {DocumentColorSpace} colorspace - The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK
	 * @return {Document} newDoc - The newly created document 
	 */
    tasks.newDocument = function(sourceDoc, colorSpace){
        var preset = new DocumentPreset();
        preset.colorMode = colorSpace;
        preset.units = sourceDoc.rulerUnits;
        
        var newDoc = app.documents.addDocument(colorSpace, preset);
        newDoc.pageOrigin = sourceDoc.pageOrigin;
        newDoc.rulerOrigin = sourceDoc.rulerOrigin;

        return newDoc;
    };
    
    /**
     * @function duplicateArtboardInNewDoc
	 * @description Takes a source document, artboardRect, and colorspace and creates a new document
	 * with the source document's units, a duplicate of the artboard and the specified color space
	 * @param {Document} sourceDoc - The source document
	 * @param {number} artboard - The artboard you want to duplicate
	 * @param {DocumentColorSpace} colorspace - The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK 
	 * @return {Document} newDoc - The newly created document
	 */
    tasks.duplicateArtboardInNewDoc = function(sourceDoc, artboard, colorspace){
        var rectToCopy = artboard.artboardRect;
        var newDoc = tasks.newDocument(sourceDoc, colorspace);
        newDoc.artboards.add(rectToCopy);	
        newDoc.artboards.remove(0);
        return newDoc;
    };

    /**
     * @function duplicateArtboardInNewDoc
	 * @description Takes a source document and colorspace and creates a new document
	 * with the source document's units and artboards and the specified color space
	 * @param {Document} sourceDoc - The source document
	 * @param {DocumentColorSpace} colorspace - The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK 
	 * @return {Document} newDoc - The newly created document
	 */
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
    
    /**
     * @function duplicateSelectionInNewDoc
	 * @description Takes a selection and a destination document and duplicates it into the destination document.
	 * Note that this does not necessarily preserve its position.
	 * @param {array} selection - An array of selected objects
	 * @param {Document} destDoc - The destination for the selected objects
	 * @return {Document} newDoc - The newly created document
	 */
    tasks.duplicateSelectionInNewDoc = function(selection, destDoc){
		var newItem = [];
		if (selection.length > 0 ) {
			for (var i = 0; i < selection.length; i++ ) {
				selection[i].selected = false;
				newItem.push(selection[i].duplicate(destDoc, ElementPlacement.PLACEATEND));
			}
		}
		return newItem;
	};

    /**
     * @function duplicateDocument
	 * @description Takes a document and desired colorspace. Creates and returns a new document in the specified colorspace with all artboards and contents duplicated
	 * @param {Document} doc - The document to duplicate
	 * @param {DocumentColorSpace} colorspace - The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK 
	 * @return {Document} newDoc - The newly created document
	 */
    //take a document and desired colorspace (e.g. DocumentColorSpace.RGB)
    //creates and returns a new document in that colorspace with all artboards and contents duplicated
	tasks.duplicateDocument = function(doc, colorspace){
        var sel = tasks.selectEverything(doc);
        var abPos = tasks.getArtboardsTopLeft(doc.artboards);
        var selPos = tasks.getCollectionTopLeft(sel);
        var offset = tasks.getOffset(selPos,abPos);
        var newDoc = tasks.duplicateArtboardsInNewDoc(doc, colorspace);
        var newSel = tasks.duplicateSelectionInNewDoc(sel,newDoc);
        var newAbPos = tasks.getArtboardsTopLeft(newDoc.artboards);
        tasks.translateSelectionTo(newSel,[newAbPos[0] + offset[0], newAbPos[1] + offset[1]]);
        return newDoc;
    };
    
    /**
     * @function newRect
	 * @description Takes left x, top y, width, and height, and returns a Rect that can be used to create an artboard
	 * @param {number} x - left position
	 * @param {number} y - top position
	 * @param {number} width - width of the rect
	 * @param {number} height - height of the rect
	 * @return {artboardRect} rect - A Rect that can be use to create an artboard
	 */    
    tasks.newRect = function(x, y, width, height) {
        var rect = [];
        rect[0] = x;
        rect[1] = -y;
        rect[2] = width + x;
        rect[3] = -(height + y);
        return rect;
    };   
    
    /*******************
    SAVING AND EXPORTING
    ********************/
    /**
     * @function scaleAndExportPNG
	 * @description Takes a document and destination file and exports a PNG at the specified scale
	 * @param {Document} doc - The document to export
	 * @param {File} destFile  - The file to export to
	 * @param {number} scaling - The percent by which to scale the PNG. For example, to scale to 50% use 50.
	 * @return {Document} newDoc - The newly created document
	 */
    tasks.scaleAndExportPNG = function(doc, destFile, scaling) {
		var options = new ExportOptionsPNG24();
		options.antiAliasing = true;
		options.transparency = true; 
		options.artBoardClipping = true;
		options.horizontalScale = scaling;
		options.verticalScale = scaling;	
		doc.exportFile(destFile, ExportType.PNG24, options);
	};


	/***
	TEXT
	****/
	
	/**
     * @function setFont
	 * @description Takes a text frame and a string with the desired font name, sets the text frame to the desired font or alerts if not found 
	 * @param {textFrame} textRef - The text frame whose font you want to set
	 * @param {String} desiredFont - A string with the name of the font you want
	 */  
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
	
	/**
     * @function createTextFrame
	 * @description Creates a text frame in a specified document with the specified message, position and font size
	 * @param {Document} doc - The document where you want to place the text frame
	 * @param {String} message - The contents for the text frame
	 * @param {array} pos - The top left corner of the text frame as an [x,y] array
	 * @param {number} size - Font size in pts
	 */  	
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
    //for future, may be better to combine the RGB and CMYK functions into one set and use color space as an argument
     
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
        for (var i = 0; i < pathItems.length; i++ ) {
            var itemColor = pathItems[i].fillColor;
            colorIndex[i] = tasks.matchToPaletteRGB(itemColor, matchArray);
        }
        return colorIndex;
    };
    
    //takes a pathItems array, startColor and endColor
    //converts all pathItems with startColor into endColor
    tasks.convertColorRGB = function(pathItems, startColor, endColor){
        for (var i = 0; i < pathItems.length; i++ ) {
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
        for (var i = 0; i < pathItems.length; i++ ) {
            var itemColor = pathItems[i].fillColor;
            colorIndex[i] = tasks.matchToPaletteCMYK(itemColor, matchArray);
        }
        return colorIndex;
    };
    
    //takes a pathItems array, startColor and endColor and converts all pathItems with startColor into endColor
    tasks.convertColorCMYK = function(pathItems, startColor, endColor){
        for (var i = 0; i < pathItems.length; i++ ) {
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
         for (var i = 0; i < pathItems.length; i++ ) {
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
        for (var i = 0; i < pathItems.length; i++ ) {
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
        for (var i = 0; i < pathItems.length; i++ ) {
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
