//#target Illustrator
//UNCOMMENT THE LINE ABOVE

/********
CSTasks is a set of functions for manipulating documents and objects in illustrator. 
It was initially designed to create variations on core objects: add elements, convert between colorspaces within a defined palette,
and export in different sizes and formats.

It uses an array [x,y] to indicate the top left corner of an object, collection, artboard, etc., to provide a consistent and
easy-to-understand frame of reference for moving objects. It also uses 2-dimensional arrays to set color palettes.

Tested using Creative Cloud 2017

Some issues include:
- Grouping and ungrouping sometimes fill in fully enclosed holes in shapes
- Color functions only consider fill color and have little error checking for objects without fill colors
- Very little error checking for locked objects and layers, or for making sure an object type is the right type
********/

var CSTasks = (function(){
 
    var CSTasks = {};    
    
    /*********************
    SELECTING AND GROUPING
    **********************/
      
    /**
     * @function selectEverything
	 * @description Selects everything in a document that is unlocked and returns an array of all the selected objects
	 * @param {Document} doc The document (should be open and active)
	 * @return {array} Array of all the selected objects
	 */
    CSTasks.selectEverything = function(doc){
        doc.selection = null;
        for (var i = 0; i < doc.pathItems.length; i++ ) {
            if (!doc.pathItems[i].layer.locked && !doc.pathItems[i].locked) doc.pathItems[i].selected = true;
        }
        return doc.selection;
    };    
    
    /**
     * @function selectContentsOnArtboard
	 * @description Selects everything on a specified artboard and returns an array of all the selected objects
	 * @param {Document} doc The document (should be open and active)
	 * @param {number} i Index of the artboard
	 * @return {array} Array of all the selected objects
	 */
    CSTasks.selectContentsOnArtboard = function(doc, i){
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
	 * @param {Document} doc The document (should be open and active)
	 * @param {Array} sel Array of selected objects
	 * @return {Group} The grouped selection
	 */

    CSTasks.groupSelection = function(doc, sel){
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
    CSTasks.ungroupOnce = function(group){
        for (var i=group.pageItems.length-1; i>=0; i--)  {
            group.pageItems[i].move(group.pageItems[i].layer, ElementPlacement.PLACEATEND);  
        }
    };
    
    /**
     * @function clearArtboard
	 * @description Deletes every unlocked object on the specified artboard
	 * @param {Document} doc The  document (should be open and active)
	 * @param {number} i Index of the artboard
	 */ 
    CSTasks.clearArtboard = function(doc, index){ //clears an artboard at the given index
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
	 * @param {artboard} artboard The Artboard object
	 * @return {array} The position of the top left corner of the artboard as an [x,y] array
	 */
    CSTasks.getArtboardTopLeft = function(artboard){ 
        var pos = [artboard.artboardRect[0], artboard.artboardRect[1]];
        return pos;
    };

	/**
     * @function getArtboardsTopLeft
	 * @description Takes an array of artboards and returns the position of their leftmost and topmost corner
	 * @param {Artboards} artboards The Artboards array 
	 * @return {array} The position of the topmost leftmost corner of all the artboards as an [x,y] array
	 */
    CSTasks.getArtboardsTopLeft = function(artboards){
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
	 * @param {array} collection The array of objects
	 * @return {array} The position of the topmost leftmost corner of all the objects as an [x,y] array
	 */
    CSTasks.getCollectionTopLeft = function(collection){
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
	 * @param {array} pos1 A position as an [x,y] array
	 * @param {array} pos2 Another position as an [x,y] array
	 * @return {array} The offset between the two positions as an [x,y] array
	 */
    CSTasks.getOffset = function(pos1, pos2){
        var offset = [pos1[0] - pos2[0], pos1[1] - pos2[1]];
        return offset;
    };

	/**
     * @function translateSelectionTo
	 * @description Takes an array of selected objects and translates their position so the top left corner
	 * is at the specified destination. Preserves relative positions of objects.
	 * @param {array} sel The array of selected objects
	 * @param {array} destination The destination position as an [x,y] array
	 */    
    CSTasks.translateSelectionTo = function(sel, destination){ 
    	var pos = CSTasks.getCollectionTopLeft(sel);
    	var offset = CSTasks.getOffset(destination, pos);
        for (var i = 0; i<sel.length; i++) {  
        	sel[i].translate(offset[0],offset[1]);
        } 
    };

	/**
     * @function translateObjectTo
	 * @description Takes an object and translates its position so the top left corner is at the specified destination
	 * @param sel The object (such as a pathItem or group)
	 * @param {array} destination The destination position as an [x,y] array
	 */ 
    CSTasks.translateObjectTo = function(object, destination){
        var offset = CSTasks.getOffset(destination, object.position);
        object.translate(offset[0],offset[1]);
    };

    /******************************************
    CREATING DOCUMENTS AND DUPLICATING CONTENTS
    ******************************************/

    /**
     * @function newDocument
	 * @description Tales a source document and colorspace and opens and returns new document with the same units in the specified colorspace.
	 * @param {Document} sourceDoc The source document
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK
	 * @return {Document} The newly created document 
	 */
    CSTasks.newDocument = function(sourceDoc, colorSpace){
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
	 * @param {Document} sourceDoc The source document
	 * @param {number} artboard The artboard you want to duplicate
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK 
	 * @return {Document} The newly created document
	 */
    CSTasks.duplicateArtboardInNewDoc = function(sourceDoc, artboard, colorspace){
        var rectToCopy = artboard.artboardRect;
        var newDoc = CSTasks.newDocument(sourceDoc, colorspace);
        newDoc.artboards.add(rectToCopy);	
        newDoc.artboards.remove(0);
        return newDoc;
    };

    /**
     * @function duplicateArtboardInNewDoc
	 * @description Takes a source document and colorspace and creates a new document
	 * with the source document's units and artboards and the specified color space
	 * @param {Document} sourceDoc The source document
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK 
	 * @return {Document} The newly created document
	 */
    CSTasks.duplicateArtboardsInNewDoc = function(sourceDoc, colorspace){
        var rectsToCopy = new Array(sourceDoc.artboards.length);
        for (var i = 0; i < sourceDoc.artboards.length; i++){
            rectsToCopy[i] = sourceDoc.artboards[i].artboardRect;
        }

        var newDoc = CSTasks.newDocument(sourceDoc, colorspace);
        
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
	 * @param {array} selection An array of selected objects
	 * @param {Document} destDoc The destination for the selected objects
	 * @return {Document} The newly created document
	 */
    CSTasks.duplicateSelectionInNewDoc = function(selection, destDoc){
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
	 * @param {Document} doc The document to duplicate
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK 
	 * @return {Document} The newly created document
	 */
    //take a document and desired colorspace (e.g. DocumentColorSpace.RGB)
    //creates and returns a new document in that colorspace with all artboards and contents duplicated
	CSTasks.duplicateDocument = function(doc, colorspace){
        var sel = CSTasks.selectEverything(doc);
        var abPos = CSTasks.getArtboardsTopLeft(doc.artboards);
        var selPos = CSTasks.getCollectionTopLeft(sel);
        var offset = CSTasks.getOffset(selPos,abPos);
        var newDoc = CSTasks.duplicateArtboardsInNewDoc(doc, colorspace);
        var newSel = CSTasks.duplicateSelectionInNewDoc(sel,newDoc);
        var newAbPos = CSTasks.getArtboardsTopLeft(newDoc.artboards);
        CSTasks.translateSelectionTo(newSel,[newAbPos[0] + offset[0], newAbPos[1] + offset[1]]);
        return newDoc;
    };
    
    /**
     * @function newRect
	 * @description Takes left x, top y, width, and height, and returns a Rect that can be used to create an artboard
	 * @param {number} x left position
	 * @param {number} y top position
	 * @param {number} width width of the rect
	 * @param {number} height height of the rect
	 * @return {artboardRect} A Rect that can be use to create an artboard
	 */    
    CSTasks.newRect = function(x, y, width, height) {
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
	 * @param {Document} doc The document to export
	 * @param {File} destFile  The file to export to
	 * @param {number} scaling The percent by which to scale the PNG. For example, to scale to 50% use 50.
	 * @return {Document} The newly created document
	 */
    CSTasks.scaleAndExportPNG = function(doc, destFile, scaling) {
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
	 * @param {textFrame} textRef The text frame whose font you want to set
	 * @param {String} desiredFont A string with the name of the font you want
	 */  
	CSTasks.setFont = function(textRef, desiredFont){
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
	 * @param {Document} doc The document where you want to place the text frame
	 * @param {String} message The contents for the text frame
	 * @param {array} pos The top left corner of the text frame as an [x,y] array
	 * @param {number} size Font size in pts
	 */  	
    CSTasks.createTextFrame = function(doc, message, pos, size){
        var textRef = doc.textFrames.add();
        textRef.contents = message;
        textRef.left = pos[0];
        textRef.top = pos[1];
        textRef.textRange.characterAttributes.size = size;
    };
    
    
    
    /************
    Color palette
    ************/
    //A quick and dirty way to deal with color palettes with corresponding RGB and CMYK values. 
    //In future, might connect to custom swatches?
    
    /**
     * @function initializeColorPalette
	 * @description Takes two equal-length arrays of corresponding RGB and CMYK color values (fairly human readable)
     * and returns an array of ColorElements (usable by the script for fill colors etc.)
	 * @param {array} RGBArray A 2-dimensional array with RGB color values (range 0-255) in the form [[R,G,B], [R2,G2,B2],...] of length n
	 * @param {array} CMYKArray A 2-dimensional array with CMYK color values (range 0-255) in the form[[C,M,Y,K],[C2,M2,Y2,K2],...] of length n
	 * @return {array} A 2-dimensional array with corresponding RGBColor and CMYKColor objects in the form [[RGBColor,CMYKColor],[RGBColor2,CMYKColor2],...] of length n
	 */ 
    CSTasks.initializeColorPalette = function(RGBArray, CMYKArray){
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
    
    /*************
    Color matching
    **************/
    
    /**
     * @function colorMatch
	 * @description Compares two colors in the specified colorspace and returns true if they are the same, false if not
	 * @param {RGBColor | CMYKColor} color1 First color to compare
	 * @param {RGBColor | CMYKColor} color2 First color to compare
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK
	 * @return {boolean} True if the colors are the same, false if not
	 */     
    CSTasks.colorMatch = function(color1, color2, colorspace){
		if (colorspace == DocumentColorSpace.RGB){
			if (Math.abs(color1.red - color2.red) < 1 &&
			Math.abs(color1.green - color2.green) < 1 &&
			Math.abs(color1.blue - color2.blue) < 1) { //can't do equality because it adds very small decimals
				return true;
			}
		}
		else if (colorspace == DocumentColorSpace.CMYK){
			if (Math.abs(color1.cyan - color2.cyan) < 1 &&
			Math.abs(color1.magenta - color2.magenta) < 1 &&
			Math.abs(color1.yellow - color2.yellow) < 1 &&
			Math.abs(color1.black - color2.black) < 1) { //can't do equality because it adds very small decimals
				return true;
			}		
		}
        return false;
    };
    
    /**
     * @function colorMatchToPalette
	 * @description Compares a color to a color palette (as created in {@link initializeColorPalette}) in the specified colorspace. 
	 * Returns the index of the matching color or -1 if no match.
	 * @param {RGBColor | CMYKColor} color  Color to compare
	 * @param {array} paletteArray Color palette as created from {@link initializeColorPalette}
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK
	 * @return {number} Index of the matching color or -1 if no match
	 */  
    CSTasks.colorMatchToPalette = function(color, paletteArray, colorspace){
    	if (colorspace == DocumentColorSpace.RGB){
			for (var i = 0; i < paletteArray.length; i++){
				if (Math.abs(color.red - paletteArray[i][0].red) < 1 &&
				Math.abs(color.green - paletteArray[i][0].green) < 1 &&
				Math.abs(color.blue - paletteArray[i][0].blue) < 1) { //can't do equality because it adds very small decimals
					return i;
				}
			}
    	}
		else if (colorspace == DocumentColorSpace.CMYK){
				for (var i = 0; i < paletteArray.length; i++){
					if (Math.abs(color.cyan - paletteArray[i][1].cyan) < 1 &&
					Math.abs(color.magenta - paletteArray[i][1].magenta) < 1 &&
					Math.abs(color.yellow - paletteArray[i][1].yellow) < 1 &&
					Math.abs(color.black - paletteArray[i][1].black) < 1) { //can't do equality because it adds very small decimals
						return i;
					}
				}	
		} 
        return -1;
    };
   
    /**
     * @function colorMatchItemsToPalette
	 * @description Compares an array of pathItems to a color palette (as created in {@link initializeColorPalette}) in the specified colorspace. 
	 * Returns an array with the index of the matching fill color (or -1 if no match) for each pathItem.
	 * @param {pathItems} pathItems  Array of pathItems
	 * @param {array} paletteArray Color palette as created from {@link initializeColorPalette}
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK
	 * @return {array} Array in which each element is the index of a pathItem's matching color (or -1 if no match)
	 */ 
    CSTasks.colorMatchItemsToPalette = function(pathItems, paletteArray, colorspace){
        var colorIndex = new Array(pathItems.length);
        for (var i = 0; i < pathItems.length; i++ ) {
            var itemColor = pathItems[i].fillColor;
            colorIndex[i] = CSTasks.colorMatchToPalette(itemColor, paletteArray, colorspace);
        }
        return colorIndex;
    };
    
    
    /***************
   	Color conversion
    ****************/
    
    /**
     * @function convertAllToColor
	 * @description Takes an array of pathItems. Converts all pathItems in unlocked layers into endColor at the specified opacity
	 * @param {pathItems} pathItems  Array of pathItems
	 * @param {RGBColor | CMYKColor} endColor  The color that you want to convert to
	 * @param {number} opcty  The percent opacity (0-100)
	 */  
    CSTasks.convertAllToColor = function(pathItems, endColor, opcty){
         for (var i = 0; i < pathItems.length; i++ ) {
        	if (!pathItems[i].layer.locked && !pathItems[i].locked){
				pathItems[i].fillColor = endColor;
				pathItems[i].opacity = opcty;
            }
        }
    };
    
    /**
     * @function convertMatchedItemsToColor
	 * @description Takes an array of pathItems. For each pathItem in an unlocked layer with a fill color that matches startColor, 
	 * converts the fill color to endColor
	 * @param {pathItems} pathItems  Array of pathItems
	 * @param {RGBColor | CMYKColor} startColor  The color that you want to convert from
	 * @param {RGBColor | CMYKColor} endColor  The color that you want to convert to
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK
	 */ 
    CSTasks.convertMatchedItemsToColor = function(pathItems, startColor, endColor, colorspace){
        for (var i = 0; i < pathItems.length; i++ ) {
            if (CSTasks.colorMatch(pathItems[i].fillColor, startColor, colorspace) && !pathItems[i].layer.locked && !pathItems[i].locked) 
            	pathItems[i].fillColor = endColor;
        }
    };

    /**
     * @function convertToPalette
	 * @description Given an array of pathItems, a color palette as created from {@link initializeColorPalette}, and an array that contains a reference to 
	 * the desired color for each pathItem, converts each pathItem to the desired color in the palette, and keeps track of colors that are not converted.
	 * Not really a standalone function, but enables converting an entire document from one colorspace to another.
	 * @param {Document} doc  document containing the items you are converting (needed to create a text frame)
	 * @param {pathItems} pathItems  Array of pathItems you wish to convert
	 * @param {array} paletteArray  Color palette as created from {@link initializeColorPalette}
	 * @param {RGBColor | CMYKColor} paletteIndex  Array in which each element corresponds to a pathItem in the pathItems array 
	 * and contains the index of one of the colors from the paletteArray.
	 * @param {DocumentColorSpace} colorspace The desired colorspace: DocumentColorSpace.RGB or DocumentColorSpace.CMYK
	 */ 
    CSTasks.convertToPalette = function(doc, pathItems, paletteArray, paletteIndex, colorspace){
        var unmatchedColors = [];
        var s = 0;
        if (colorspace == DocumentColorSpace.CMYK) s = 1;
        for (var i = 0; i < pathItems.length; i++ ) {
            if (paletteIndex[i] >=0 && paletteIndex[i] < paletteArray.length) pathItems[i].fillColor = paletteArray[paletteIndex[i]][s];
            else {
                var unmatchedColor = "(" + pathItems[i].fillColor.red + ", " + pathItems[i].fillColor.green + ", " + pathItems[i].fillColor.blue + ")";
                unmatchedColors.push(unmatchedColor);	
            }
        }
        if (unmatchedColors.length > 0){  
        	alert("One or more colors don't match the brand palette and weren't converted.");        
            unmatchedColors = CSTasks.getUniqueElements(unmatchedColors);
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
        	
            CSTasks.createTextFrame(doc, unmatchedString, errorMsgPos, 18);
        }
    };
    
    /**
     * @function getUniqueElements
	 * @description Takes an array and returns an array with the unique elements in sorted order
	 * @param {array} a  Starting array
	 * @param {array} uniq Array with the sorted unique elements
	 */ 
    CSTasks.getUniqueElements = function(a) {
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

    return CSTasks;
}());