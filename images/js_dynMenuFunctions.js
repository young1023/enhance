<!--BEGIN js_dynMenuFunctions.js-->
//--------------------------------------------------------------
//This code checks to see if push is defined and if it returns
//the last value pushed.  If so, set push definition to null.
//In JavaScript1.2, it returns the value pushed instead of 
//returning the new length of the array (as in 1.0, 1.3).
//This is a moot point while we use 1.0.  It's here in case
//we switch to a different JavaScript version later.
//--------------------------------------------------------------
if (Array.prototype.push && ([0].push(true)==true))
        Array.prototype.push = null;

//If no push function exists, create it!  Will allow you to push
//multiple arguments in a single call to push.
if(!Array.prototype.push) {

    function array_push() {
        for(i=0;i<arguments.length;i++){
            this[this.length] = arguments[i];
        }
        return this.length;
    }
    Array.prototype.push = array_push;
}
//If no pop function exists, create it!
if(!Array.prototype.pop) {

    function array_pop(){
        lastElement = this[this.length-1];
        this.length = Math.max(this.length-1,0);
        return lastElement;
    }

    Array.prototype.pop = array_pop;
}
var gstrBrowserType = null
if( document.layers ) {
	gstrBrowserType="Netscape";
	function reDo(){ window.location.reload() }
	function setResize(){setTimeout("window.onresize=reDo",500);}
	window.onload = setResize;
}
else if( document.all ) {
	gstrBrowserType="IE";
}

//---------------------------------------------------------------------------
// setMenuVisibility - will change the state of a menu to the value passed in.
//                     Netscape=[hide,show] IE=[hidden,visible]
// Arguments: strMenuItemDiv  - the actual menu item div as a string
//            strNewMenuState - the state to change the div to
//---------------------------------------------------------------------------
function setMenuVisibility( strMenuDiv, strNewMenuState ){
	var objMenuDiv;
	if( gstrBrowserType == "Netscape")
	{
		objMenuDiv = eval(strMenuDiv);
  		if( objMenuDiv ) objMenuDiv.visibility = strNewMenuState;
	}
	else if( gstrBrowserType == "IE" )
	{
		objMenuDiv=eval(strMenuDiv);
		if( objMenuDiv ) objMenuDiv.style.visibility = strNewMenuState;
	}
	else if( gstrBrowserType == "NS6" )
	{
		objMenuDiv=eval("document.getElementById(\"" + strMenuDiv + "\")");
		if( objMenuDiv ){
		objMenuDiv.style.visibility = strNewMenuState;
		}
	}	
	else
	{
		return
	}
}
//---------------------------------------------------------------------------
// setMenuItemBgColor - will change the background color of a menu item to
//                      the color passed in.  This provides the highlighting
//                      effect when the mouse is over a particular menu item.
// Arguments: strMenuItemDiv - the actual menu item div as a string
//            strColor       - the color to change to
//---------------------------------------------------------------------------
function setMenuItemBgColor( strMenuItemDiv, strColor ){
	var objMenuItemDiv;

	if( gstrBrowserType == "Netscape")
	{
		objMenuItemDiv = eval(strMenuItemDiv);
  		if( objMenuItemDiv ) objMenuItemDiv.bgColor = strColor;
	}
	else if( gstrBrowserType=="IE" )
	{
		objMenuItemDiv=eval(strMenuItemDiv);
		if( objMenuItemDiv ) objMenuItemDiv.style.backgroundColor = strColor;
	}
	else if( gstrBrowserType=="NS6" )
	{
		objMenuItemDiv=eval("document.getElementById(\"" + strMenuItemDiv + "\")");
		if( objMenuItemDiv ) objMenuItemDiv.style.backgroundColor = strColor;
	}	
	else
	{
		return
	}
}
//---------------------------------------------------------------------------
// createNetscapeMenu - create the layers for Netscape browsers.  In order to
//	get a border in Netscape, I had to add an extra layer.  The layer was
//	required because we have a left/right margin on the text layers that
//	show the color of the layer behind the text layer.  This "layer behind"
//	was the main layer which was the color of the border.  Thus, the color
//	of the left/right margin all the way out to the border was the color of
//	the border instead of the color of the menu item.
//
//	I put an extra layer in.  The main layer is the color of the desired 
//	border.  The second layer is inset by one pixel with the color of the
//	menu items.  When we apply the margins to the menu items, now you see
//	the menu item color of layer 2 behind the menu item instead of the 
//	border color of the main layer.  Sorry for the novel-like comments!
//
// Arguments: strmenuName - name of the menu to be created. This is used to
//                          index into garyMenuData
//            intLeftPos  - value of left starting point of the menu in pixels
//            intWidth    - value of width of menu in pixels
//            strAlignText- the alignment of the text in the menu (left unless
//                          you are the last menu, then right).
//---------------------------------------------------------------------------
function createNetscapeMenu( strMenuName, intLeftPos, intWidth, strTextAlign ){
	var strLayerDef = "";                    //The HTML definition of the whole menu
	var strID = "";                          //The id of the menu items - it is the first 3 letters
                                                 //of the menu name followed by a counter number (0,1,2...)
	var intLineHeight = 16;                  //The height of the menu item lines
	var intInsideWidth = intWidth - 2;       //The width of layer 2
	var intItemWidth = intInsideWidth - 2;   //The width of text items which can be highlighted
	var intMargin = 2;                       //The width of the left/right margins
	//var separator="<img src=\"/common/assets/images/bar.gif\" width=\"" + (intWidth-20) + "\" height=\"1\">"
	var isMac = (navigator.appVersion.indexOf("Mac") != -1);	
	var intTopPos = 102;                     //The absolute top of all the menu
	var strLayerStyle = "gNetMenuItemStyle"  //string designating the layer style, specifies either PC or Mac style
	if( isMac ) strLayerStyle = "gNetMenuItemStyleMac"

	if( strTextAlign == "right" ) strTextAlign = " align=\"right\"";
	else strTextAlign = "";

	//The first layer holds the second layer (which holds all the other elements).  The only visible portion of it is 1px around for the border.
	strLayerDef  = "<layer id=\""+ strMenuName +"\" visibility=\"hide\" left=\"" + intLeftPos + "\" top=\"" + intTopPos + "\" "
 	strLayerDef += "width=\"" + intWidth + "\" z-index=\"1\" bgcolor=\"" + gstrBorderColor + "\" onmouseover=\"showMenu('" + strMenuName + "')\" onmouseout=\"hideMenu('" + strMenuName + "')\">"

	//The second layer just holds all the other elements.  It's background color is the same as that of the menu items so when the menu
	//items are cropped, we don't see the color of the border on the area of the element left transparent by the crop.
	strLayerDef += "<layer id=\""+ strMenuName +"1\" left=\"1\" top=\"1\" "
 	strLayerDef += "width=\"" + intInsideWidth + "\" z-index=\"7\" bgcolor=\"" + gstrMenuColor + "\">"

	var intTopCoord = 1
	strLayerDef += "<layer id=\"firstspacer\" bgcolor=\"" + gstrMenuColor +"\" left=\"1\" width=\"" + intItemWidth + "\" height=\"3\" z-index=\"7\" top=\"" + intTopCoord + "\"></layer>"
	intTopCoord += 2;

	for( var intItem=0; intItem < garyMenuData[strMenuName].length; intItem++ ) {
		intTopCoord += 0;
		strID = strMenuName.substr(0,3) + intItem
		strLayerDef += "<layer id=\"" + strID + "\" clip=\"" + intMargin + ",0," + (intWidth-intMargin) + "," + intLineHeight + "\" class=\"" + strLayerStyle + "\" "
		strLayerDef += "z-index=\"7\" width=\"" + intItemWidth + "\" top=\"" + intTopCoord + "\" bgcolor=\"" + gstrMenuColor + "\" "
		strLayerDef += "onmouseover=\"setMenuItemBgColor('document.layers[\\'" + strMenuName + "\\'].document.layers[\\'" + strMenuName + "1\\'].document.layers[\\'" + strID + "\\']', '" + gstrMenuColorOn + "')\" "
		strLayerDef += "onmouseout=\"setMenuItemBgColor('document.layers[\\'" + strMenuName + "\\'].document.layers[\\'" + strMenuName + "1\\'].document.layers[\\'" + strID + "\\']', '" + gstrMenuColor + "')\"> "
		strLayerDef += "<p" + strTextAlign + "><a class=\"" + style + "\" href=\"" + garyMenuData[strMenuName][intItem][1] + "\"> "
		strLayerDef += garyMenuData[strMenuName][intItem][0] + "</a></p></layer>"
		intTopCoord += intLineHeight;
	}
	intTopCoord += 7;
	strLayerDef += "<layer id=\"lastspacer1\" bgcolor=\"" + gstrMenuColor + "\" width=\"" + intItemWidth + "\" height=\"1\" z-index=\"7\" top=\"" + intTopCoord + "\"></layer>"
	intTopCoord += 1;
	strLayerDef += "<layer id=\"lastspacer\" bgcolor=\"" + gstrBorderColor + "\" width=\"" + intInsideWidth + "\" height=\"1\" z-index=\"7\" top=\"" + intTopCoord + "\"></layer>"
	strLayerDef += "</layer>" //end the 2nd layer
	strLayerDef += "</layer>" //end the main layer for the whole menu

	document.write( strLayerDef );
}
//---------------------------------------------------------------------------
// createIEMenu - create the layers for Internet Explorer browsers
//
// Arguments: strmenuName - name of the menu to be created. This is used to
//                          index into garyMenuData
//            intLeftPos  - value of left starting point of the menu in pixels
//            intWidth    - value of width of menu in pixels
//            strAlignText- the alignment of the text in the menu (left unless
//                          you are the last menu, then right.
//---------------------------------------------------------------------------
function createIEMenu( strMenuName, intLeftPos, intWidth, strTextAlign, style ){

	var strDivDef = "";                     //The HTML definition of the whole menu
	var strID     = "";                     //The id of the menu items - it is the first 3 letters
	var intMargin = 5;                    //The width of the left/right margins
	var isMac = (navigator.appVersion.indexOf("Mac") != -1);
	if( style.search(/AOL/i) >= 0 ){
		var intTopPos = 102;
	}
	else{
		var intTopPos = 100;                    //The absolute top of all the menus
	}
	var strLayerStyle = style  //string designating the layer style, specifies either PC or Mac style
	if( isMac ) {
		strLayerStyle = style + "Mac";
		//intTopPos = 115;    
		//intLeftPos = intLeftPos + 8;               
	}

	//var separator="<img src=\"/common/assets/images/bar.gif\" width=\"" + intWidth + "\" height=\"1\">"
	strDivDef  = "<div id=\"" + strMenuName + "\" style=\"visibility:'hidden'; position:'absolute'; left:'" + intLeftPos + "'; top:'" + intTopPos + "'; cursor: hand;"
	strDivDef += "width:'" + intWidth + "px'; z-index:3; border:1px solid " + gstrBorderColor + "; background-color:'" + gstrMenuColor + "';\" onmouseover=\"showMenu('" + strMenuName + "')\" "
	strDivDef += "onmouseout=\"hideMenu('" + strMenuName + "')\">"
	strDivDef += "<div id=\"separator\" style=\"position:'relative'; line-height:4px; width:" + (intWidth-intMargin) + "px; background-color:'"
	strDivDef += gstrMenuColor + "';\"><span style=\"line-height:" + intMargin + "px;\">&nbsp;</span></div>"

	for( var intItem=0; intItem < garyMenuData[strMenuName].length; intItem++ ) {
		
		//strDivDef += "<div id=\"separator\" style=\"position:'relative'; width:" + (intWidth-intMargin) + "px; margin-right:" + intMargin + "px; margin-left:" + intMargin + "px; background-color:'"
		//strDivDef += gstrMenuColor + "';\"><center>" + separator + "</center></div>"
		strID = strMenuName.substr(0,3) + intItem;
		strDivDef += "<div id=\"" + strID + "\" style=\"position:'relative'; text-align:'" + strTextAlign + "'; "
		strDivDef += "background-color:'" + gstrMenuColor + "'; margin-left:" + (intMargin+3) + "px; margin-right:" + (intMargin+3) + "px; width:" + (intWidth-6) + "px;\" "
		strDivDef += "onmouseover=\"setMenuItemBgColor('document.all[\\'" + strID + "\\']','" + gstrMenuColorOn + "')\" "
		strDivDef += "onmouseout=\"setMenuItemBgColor('document.all[\\'" + strID + "\\']','" + gstrMenuColor + "')\" "
		strDivDef += "onClick=\"location.href='" + garyMenuData[strMenuName][intItem][1] + "'\"><span class=\"" + strLayerStyle + "\" >" + garyMenuData[strMenuName][intItem][0]
		strDivDef += "</span></div>"
	}
	//strDivDef += "<div id=\"separator\" style=\"position:'relative'; margin-left:" + intMargin + "px; margin-right:" + intMargin + "px; width:" + (intWidth-intMargin) + "px; background-color:'"
	//strDivDef += gstrMenuColor + "';\"><center>" + separator + "</center></div>"

	strDivDef += "<div id=\"lastseparator\" style=\"position:'relative';width:'" + intWidth + "px';background-color:'" + gstrMenuColor + "';\"><span style=\"line-height:" + intMargin + "px;\">&nbsp;</span></div>";
	strDivDef += "</div>"
	document.write( strDivDef );
}

//---------------------------------------------------------------------------
// createNS6Menu - create the layers for Netscape 6 browsers
//
// Arguments: strmenuName - name of the menu to be created. This is used to
//                          index into garyMenuData
//            intLeftPos  - value of left starting point of the menu in pixels
//            intWidth    - value of width of menu in pixels
//            strAlignText- the alignment of the text in the menu (left unless
//                          you are the last menu, then right.
//---------------------------------------------------------------------------
function createNSsixMenu( strMenuName, intLeftPos, intWidth, strTextAlign, style ){

	var strDivDef = "";                     //The HTML definition of the whole menu
	var strID     = "";                     //The id of the menu items - it is the first 3 letters
	var intMargin = 5;      
	//intWidth = intWidth + 10;              //The width of the left/right margins
	var isMac = (navigator.appVersion.indexOf("Mac") != -1);
	//if( style.search(/AOL/i) >= 0 ){
	//	var intTopPos = 106;
	//}
	//else{
		var intTopPos = 102;                    //The absolute top of all the menus
	//}
	
	
	var strLayerStyle = style  //string designating the layer style, specifies either PC or Mac style
	if( isMac ) {
		strLayerStyle = style + "Mac";
		//intTopPos = 115;    
		//intLeftPos = intLeftPos + 8;               
	}

	//var separator="<img src=\"/common/assets/images/bar.gif\" width=\"" + intWidth + "\" height=\"1\">"
	strDivDef  = "<div id=\"" + strMenuName + "\" style=\"visibility:hidden;position:absolute;left:" + intLeftPos + ";top:" + intTopPos + ";cursor:pointer;"
	strDivDef += "width:" + intWidth + "px;z-index:3;border:1px solid " + gstrBorderColor + ";background-color:" + gstrMenuColor + ";\" onmouseover=\"showMenu('" + strMenuName + "')\" "
	strDivDef += "onmouseout=\"hideMenu('" + strMenuName + "')\">"
	strDivDef += "<div id=\"separator\" style=\"position:relative; line-height:4px; width:2px; background-color:"
	strDivDef += gstrMenuColor + ";\"><span style=line-height:1px;>&nbsp;</span></div>"

	for( var intItem=0; intItem < garyMenuData[strMenuName].length; intItem++ ) {

		//strDivDef += "<div id=\"separator\" style=\"position:'relative'; width:" + (intWidth-intMargin) + "px; margin-right:" + intMargin + "px; margin-left:" + intMargin + "px; background-color:'"
		//strDivDef += gstrMenuColor + "';\"><center>" + separator + "</center></div>"
		strID = strMenuName.substr(0,3) + intItem;
		strDivDef += "<div id=\"" + strID + "\" style=\"position:relative; text-align:" + strTextAlign + "; "
		strDivDef += "background-color:" + gstrMenuColor + "; margin-left:" + (intMargin+1) + "px; margin-right:" + (intMargin+3) + "px; width:" + (intWidth-6) + "px;\" "
		strDivDef += "onmouseover=\"setMenuItemBgColor('" + strID + "','" + gstrMenuColorOn + "')\" "
		strDivDef += "onmouseout=\"setMenuItemBgColor('" + strID + "','" + gstrMenuColor + "')\" "
		strDivDef += "onClick=\"location.href='" + garyMenuData[strMenuName][intItem][1] + "'\"><span class=\"" + strLayerStyle + "\" >" + garyMenuData[strMenuName][intItem][0]
		strDivDef += "</span></div>"
	}
	//strDivDef += "<div id=\"separator\" style=\"position:'relative'; margin-left:" + intMargin + "px; margin-right:" + intMargin + "px; width:" + (intWidth-intMargin) + "px; background-color:'"
	//strDivDef += gstrMenuColor + "';\"><center>" + separator + "</center></div>"

	strDivDef += "<div id=\"separator\" style=\"position:relative; line-height:4px; width:6px; background-color:"
	strDivDef += gstrMenuColor + ";\">&nbsp;</div>"

	strDivDef += "</div>"
	document.write( strDivDef );

}

//---------------------------------------------------------------------------
// showMenu - This function determines which browser you are using and sends
//            the appropriate (browser-specific) CSS string along to the 
//            setMenuVisibility function to make the current menu visible.
//
// Arguments: strMenu - the name of the menu to show
//---------------------------------------------------------------------------
function showMenu( strMenu ) {
	if( gstrBrowserType == "Netscape" ) {
		setMenuVisibility("document.layers[\"" + strMenu + "\"]","show")
	}
	else if( gstrBrowserType == "IE" ) {
		setMenuVisibility("document.all[\"" + strMenu + "\"]","visible")
	}
	else if( gstrBrowserType == "NS6" ) {
		setMenuVisibility(strMenu,'visible')
	}
}
//---------------------------------------------------------------------------
// hideMenu - This function determines which browser you are using and sends
//            the appropriate (browser-specific) CSS string along to the 
//            setMenuVisibility function to hide the current menu.
//
// Arguments: strMenu - the name of the menu to hide
//---------------------------------------------------------------------------
function hideMenu( strMenu ) {
	if( gstrBrowserType == "Netscape" ) {
		setMenuVisibility("document.layers[\"" + strMenu + "\"]","hide")
	}
	else if( gstrBrowserType == "IE" ) {
		setMenuVisibility(document.all[strMenu],"hidden")
	}
	else if( gstrBrowserType == "NS6" ) {
		setMenuVisibility(strMenu,'hidden')
	}	
}
<!--END   js_dynMenuFunctions.js-->