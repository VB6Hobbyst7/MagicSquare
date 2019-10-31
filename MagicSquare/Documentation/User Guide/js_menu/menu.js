
	var openImg = new Image();
	var closedImg = new Image();
	var fileImg = new Image(); 
	
	var intX, intY;
	var selectedMenu;
// JScript source code
	//constructors in JavaScript
	/*
	The keyword "this" indicates that properties are being assigned to the objects 
	created by the constructor. The menuItem() constructor code requires two 
	parameters for each menuItem object you create: some text and a link. The 
	constructor assigns the parameters to properties of the menuItem named "text" 
	and "link" respectively. Later, when you create a menuItem object, you'll be 
	able to reference the text of the menuItem by using the familiar dot (.) 
	notation.
	uUse: var myItem = new menuItem('DevX','http://www.devx.com');
			alert(myItem.text);
			---------------------
			var myItem = new menuItem();
			myItem.text = 'DevX';
			myItem.link = 'http://www.devx.com';
	*/
   function menuItem(text, link){
      this.text = text;
      this.link = link;
   }
   
   function menuTrigger(name, text){
      this.name = name;
      this.text = text;
   }
   //create menu objects
   /*
	The itemArray variable holds the collection of menuItems for a menu object. The 
	collection itself is nothing more than a plain old JavaScript array. The args 
	variable represents an array of all the parameters passed into the menu() 
	constructor. The constructor assigns the first parameter to the name property 
	and the second parameter to the trigger property. It's important to note that 
	the trigger property will actually hold a menuTrigger object. Next, the code 
	loops through the arguments assigning each menuItem for this menu to the 
	itemArray. The first two parameters have already been used, therefore the loop 
	starts at 2 , not 0. You must subtract 2 from the value of i to properly index 
	the elements in itemArray. Finally, the code assigns the itemArray itself to the 
	menuItems property.
   */
   function menu(){
      var itemArray = new Array();
      var args = menu.arguments;
      this.name = args[0];
      this.trigger = args[1];
      for(i=2; i<args.length; i++){
         itemArray[i-2] = args[i];
      }
      this.menuItems = itemArray;
      this.write = writeMenuInDocument;		//handle the display of the menu
      this.position = positionMenu; //handle positioning of the menu
    }
   //handle positioning of the menu
   /*
		adds three additional properties to each menu object: top, left, and width
   */
   function positionMenu(top,left,width){
      this.top = top;
      this.left = left;
      this.width = width;
   }
   /*
	build a string to write to the document. 
	First, it builds a <div> element for the menuTrigger object for each menu. 
	Next, it builds a <div> element for all the menuItems contained in the menuItems array 
	within the menu object.
   */
   
   function writeMenuInDocument()
   {
		intX = 0;
		intY = 0;
		var menuText = writeMenu(this);
		menuText = '<tr><td>' + menuText + '</td></tr>';
		var SText = menuText;	
		//alert(SText);	
		document.write(menuText);		
		document.close();		
		//WriteToFile(SText);
   }
   
   function writeMenu(objMenu){	   	
      var menuText = "";
      menuText += '<table border="0" cellspacing="0" cellpadding="0"><tr><td>';      
      menuText += '<table border="0" cellspacing="0" cellpadding="1">';      
      menuText += '<tr><td>'; 	           
      menuText += '<img align="right" onClick="showMenuAndUpdateStyle(\'' + objMenu.name + '\',\'td' + objMenu.trigger.name + '\')" src="' + closedImg.src + '" id="I'+ objMenu.name +'" class="imgLink" width="16" height="16" />';
      menuText += '</td><td id="td' + objMenu.trigger.name + '" class="menuText">';     
      menuText += '<a onClick="showMenuAndUpdateStyle(\'' + objMenu.name + '\',\'td' + objMenu.trigger.name + '\')">'      
      menuText +=  objMenu.trigger.text ; 
      menuText += '</a></td></tr>';
      menuText += '</table></td></tr>';
      menuText += '<tr><td><div id="' + objMenu.name + '" ';     
      menuText += 'class="menu" style="visibility:hidden;display:none;';
      if (objMenu.top > -1)
      {
		menuText += 'position:absolute;top: ' + (objMenu.top+23);		
	  } 
	  else
	  {
		menuText += 'position:relative;';
	  }	  
	  if (objMenu.left > -1)
	  {
		menuText += '; left: ' + (objMenu.left);			
	  }       
      if (objMenu.width > -1)
      {
		menuText += ';width: ' + objMenu.width + ';';		    
      }        
      menuText += '" >';
      intX = 0;
      intY = 0;
      menuText += '<table border="0" width="' + objMenu.width + '">';      
      var i;
      for(i=0; i<objMenu.menuItems.length; i++){		
         menuText += writeItem(objMenu.menuItems[i]);        
      }
      menuText += '</table>';
      menuText += '</div></td></tr></table>';  
      return menuText;    
   }
   
   function writeItem(objMenuItem)
   {
		var menuText = "";
		if (objMenuItem instanceof menuItem)
		{	
			menuText += '<tr><td>';
			menuText += '<table border="0" cellspacing=0 cellpadding=0>';
			menuText += '<tr><td width="22" align="right"><img class="imgLink" onClick=\'redirectAndUpdateStyle("' + objMenuItem.link + '","' + objMenuItem.text.replace(/ /g, "") + '")\' src="' + fileImg.src + '" width="16" height="16"/></td>';
			
			menuText += '<td id="' + objMenuItem.text.replace(/ /g, "") + '" class="menuText">';
			menuText += '<a onClick=\'redirectAndUpdateStyle("' + objMenuItem.link + '","' + objMenuItem.text.replace(/ /g, "") + '")\'>';			
			menuText += objMenuItem.text + '</a></td></tr></Table>';
			menuText += '</td></tr>';
			intMultiple = objMenuItem.text.length;			
			while (intMultiple > 0)
			{
				intX += 18;	
				intMultiple -=20;
			}		
		}
		else
		{
			if (objMenuItem instanceof menu)
			{							
				objMenuItem.position(-1,10,230);
				menuText += '<tr><td>';				
				menuText += writeMenu(objMenuItem);	
				menuText += '</td></tr>';			
			}
		}		
        return menuText;
   }  
   
   
   function showMenuAndUpdateStyle(menu, strMenu){  
	//Changes the style of the menu items.
	changeStyle(strMenu);
	
	//Shows/hides the menu.
	var objStyle = document.getElementById(menu).style;	  
	if(objStyle.display=="block")
		objStyle.display="none";
	else
		objStyle.display="block"; 

	if (objStyle.visibility == 'hidden')		
		objStyle.visibility = 'visible';
	else
		objStyle.visibility = 'hidden';     
    swapFolder("I" + menu);
   }
   
   function showMenu(menu, strMenu){
	//Shows the menu.
	var objStyle = document.getElementById(menu).style;	  
	objStyle.display="block";	
	objStyle.visibility = 'visible';
	
	var objImg = document.getElementById('I' + menu);
	if(objImg.src.indexOf('folderClosed.gif')>-1)
		objImg.src = openImg.src;
	}
   
   function hideMenu(menu){
   alert(menu);
      if(mnuSelected!='')
         document.getElementById(menu).style.visibility = 'hidden';
   }
   
   function swapFolder(img){
	objImg = document.getElementById(img);
	if(objImg.src.indexOf('folderClosed.gif')>-1)
		objImg.src = openImg.src;
	else
		objImg.src = closedImg.src;
	}

	function redirectAndUpdateStyle(strLink, strText)
	{
		//parent.frames[1].location = strLink;
		parent.top.location = strLink;
		changeStyle(strText);	
	}
	
	/*function WriteToFile(sText)
	{
		var fso = new ActiveXObject("Scripting.FileSystemObject");
		var FileObject = fso.CreateTextFile("C:\\xxxx.txt", true); // 8=append, true=create if not exist, 0 = ASCII
		FileObject.write(sText);
		FileObject.close();
	}*/
	


function changeStyle(menu)
{
	if (selectedMenu != null)
	{
		var objSelectedMenuStyle = document.getElementById(selectedMenu).style; 	
		objSelectedMenuStyle.background="white";
		objSelectedMenuStyle.color="black";		
	}
	
	var objNewSelectedMenuStyle = document.getElementById(menu).style; 	
	objNewSelectedMenuStyle.background="#23759D";
	objNewSelectedMenuStyle.color="white";

	if (findPosY(document.getElementById(menu)) + 15 > document.body.clientHeight + document.body.scrollTop)
	{
		document.body.scrollTop = findPosY(document.getElementById(menu)) + 15 -document.body.clientHeight;
	}
	selectedMenu = menu;
}

function findPosY(obj){
	var curtop = 0;
	if (obj.offsetParent){
		while (obj.offsetParent){
			curtop += obj.offsetTop;
			obj = obj.offsetParent;
		}
	}else if (obj.y)
		curtop += obj.y;
	return curtop;
}

   function initLeftMenu(folderParentImages)
   {
	  openImg.src = folderParentImages + "images/menu/folderOpen.gif";
	  openImg.width = "16";
	  openImg.height = "16";
	  closedImg.src = folderParentImages + "images/menu/folderClosed.gif";
	  closedImg.width = "16";
	  closedImg.height = "16";
	  fileImg.src = folderParentImages + "images/menu/file.gif";
	  fileImg.width = "16";
	  fileImg.height = "16";
   
      var iText = "<table border=0 cellspacing=0 cellpadding=0>";
	  document.write(iText);		
	  document.close();		
   }
   
   function selectMenuItem(menuItemTitle)
   {
   }
   
    
   function fin()
   {
	   var fText = "</table>";
	  document.write(fText);		
	  document.close();		
   }
	





