// JScript source code
	var opt1 = "C#.NET";
	var opt2 = "ASP";
	var opt3 = "C++";
	var opt4 = "C++.NET";
	var opt5 = "Coldfusion";
	var opt6 = "J#.NET";
	var opt7 = "Java";
	var opt8 = "PHP";
	var opt9 = "VB.NET";
	var opt10 = "VB6";
	var opt11 = "VBS";	
	
	var selectedImg = new Image();
	var inactiveImg = new Image();
	
	// This variable is used to select the API's type: .Net or Java
	var isAPIForDotNet = "true";
	var isOnlineVersion = "true"; 
	
	function changeFace(obj, option){
		if(obj!=null){
			if(option == 1){
				obj.style.color = 'gray';
				obj.style.cursor = 'pointer';
				obj.style.textDecoration = 'underline';
			}
			else{
				obj.style.color = '#4d614b';
				obj.style.textDecoration = 'none';
			}
		}
	}
	
	function selectTab(chosen, codeSectionNumber){
	    
		if (codeSectionNumber > 1)	
	    {
			codeSectionNumber = '_' + codeSectionNumber;
		}
		else{
			codeSectionNumber = '';
		}

		//change code
		document.getElementById('Code' + codeSectionNumber).innerHTML = document.getElementById(chosen + codeSectionNumber).innerHTML;	
						
		//change the image
		if (document.getElementById('td_' + opt1 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt1 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		if (document.getElementById('td_' + opt1 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt1 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt2 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt2 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt3 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt3 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt4 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt4 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt5 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt5 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt6 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt6 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt7 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt7 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt8 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt8 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt9 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt9 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt10 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt10 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		
		if (document.getElementById('td_' + opt11 + codeSectionNumber)!=null){
			document.getElementById('td_' + opt11 + codeSectionNumber).style.backgroundImage = 'url(../images/inactive.gif)';
		}
		document.getElementById('td_' + chosen + codeSectionNumber).style.backgroundImage = 'url(../images/selected.gif)';	
	}

	function toggleLayer(layerid)
	{
	    var style2;
		if (document.getElementById){
			// this is the way the standards work				
			style2 = document.getElementById(layerid).style;				
			style2.display = style2.display? "":"block";
		}
		else if (document.all){
			// this is the way old msie versions work
			style2 = document.all[layerid].style;				
			style2.display = style2.display? "":"block";
		}
		else if (document.layers){
			// this is the way nn4 works				
			style2 = document.layers[layerid].style;
			style2.display = style2.display? "":"block";
		}
	}
		
	function callAPIMethod(method)
	{		
		if (isAPIForDotNet == "true")
			parent.location = method;
		else
			parent.location = replaceParameters(method);
	}
	
	function changeMethodsName()
	{
		if (isAPIForDotNet == "false")
		{			
			if (document.getElementById("spanDataSet"))
				document.getElementById("spanDataSet").innerHTML = document.getElementById("spanDataSet").innerHTML.replace("DataSet","ResultSet");
			if (document.getElementById("spanDataSet2"))
				document.getElementById("spanDataSet2").innerHTML = document.getElementById("spanDataSet2").innerHTML.replace("DataSet","ResultSet");
			if (document.getElementById("spanDataSet3"))
				document.getElementById("spanDataSet3").innerHTML = document.getElementById("spanDataSet3").innerHTML.replace("DataSet","ResultSet");
			if (document.getElementById("spanDataSet4"))
				document.getElementById("spanDataSet4").innerHTML = document.getElementById("spanDataSet4").innerHTML.replace("DataSet","ResultSet");
			if (document.getElementById("spanDataSet5"))
				document.getElementById("spanDataSet5").innerHTML = document.getElementById("spanDataSet5").innerHTML.replace("DataSet","ResultSet");
			if (document.getElementById("spanDataSet6"))
				document.getElementById("spanDataSet6").innerHTML = document.getElementById("spanDataSet6").innerHTML.replace("DataSet","ResultSet");		
			if (document.getElementById("spanDataSet7"))
				document.getElementById("spanDataSet7").innerHTML = document.getElementById("spanDataSet7").innerHTML.replace("DataSet","ResultSet");
			if (document.getElementById("spanDataSet8"))
				document.getElementById("spanDataSet8").innerHTML = document.getElementById("spanDataSet8").innerHTML.replace("DataSet","ResultSet");	
			if (document.getElementById("spanDataSet9"))
				document.getElementById("spanDataSet9").innerHTML = document.getElementById("spanDataSet9").innerHTML.replace("DataSet","ResultSet");	
			if (document.getElementById("spanDataSet10"))
				document.getElementById("spanDataSet10").innerHTML = document.getElementById("spanDataSet10").innerHTML.replace("DataSet","ResultSet");				
		}
	}
	
	function replaceParameters(methodName)
	{
		methodName = methodName.replace("System.Data.DataSet", "java.sql.ResultSet");
		methodName = methodName.replace("ReadXLSSheet_AsXML(System.IO.Stream, System.IO.Stream", "ReadXLSSheet_AsXML(java.io.OutputStream, java.io.InputStream");
		methodName = methodName.replace("ReadXLSXSheet_AsXML(System.IO.Stream, System.IO.Stream", "ReadXLSXSheet_AsXML(java.io.OutputStream, java.io.InputStream");
		methodName = methodName.replace("ReadXLSBSheet_AsXML(System.IO.Stream, System.IO.Stream", "ReadXLSBSheet_AsXML(java.io.OutputStream, java.io.InputStream");
		methodName = methodName.replace("ReadExcelWorksheet_AsXML(System.IO.Stream", "ReadExcelWorksheet_AsXML(java.io.OutputStream");
		methodName = methodName.replace("FromDataSet(System.IO.Stream", "FromResultSet(java.io.OutputStream");
		methodName = methodName.replace("FromList(System.IO.Stream", "FromList(java.io.OutputStream");
		methodName = methodName.replace("WriteXLSFile(System.IO.Stream", "WriteXLSFile(java.io.OutputStream");
		methodName = methodName.replace("WriteXLSXFile(System.IO.Stream", "WriteXLSXFile(java.io.OutputStream");
		methodName = methodName.replace("WriteXLSBFile(System.IO.Stream", "WriteXLSBFile(java.io.OutputStream");
		methodName = methodName.replace("WriteTXTFile(System.IO.Stream", "WriteTXTFile(java.io.OutputStream");
		methodName = methodName.replace("WriteCSVFile(System.IO.Stream", "WriteCSVFile(java.io.OutputStream");
		methodName = methodName.replace("WriteHTMLFile(System.IO.Stream", "WriteHTMLFile(java.io.OutputStream");
		methodName = methodName.replace("WriteXMLFile(System.IO.Stream", "WriteXMLFile(java.io.OutputStream");
		methodName = methodName.replace("DataSet", "ResultSet");
		methodName = methodName.replace(/bool/g,"boolean").replace(/System.String/g, "java.lang.String");
		methodName = methodName.replace("System.IO.Stream", "java.io.InputStream"); /*covers all read and load methods, write methods are cover above separately*/
		methodName = methodName.replace("System.Drawing.Color", "java.awt.Color");
		return methodName;
	}
	
	/* Footer min 148,max 155*/
var  cons = 153;
var  pDiff = cons;
function  SetFooter()
{
    var divFooter = document.getElementById('divFooter');
    if (divFooter == null) { return; }
    var  ffClientHeight;
    var ffFooterWidth;
    if  (ffClientHeight == undefined || ffClientHeight == null)
        ffClientHeight = document.body.parentNode.clientHeight;

    if  (ffClientHeight < document.body.clientHeight + pDiff)
    {
        divFooter.style.position = "relative";
        pDiff = 0;
        divFooter.style.width = "100px";
    }
    else
     {
        divFooter.style.position = "absolute";
        divFooter.style.bottom = "0px";
        divFooter.style.width = "";
        pDiff = cons;
    }
    
    ffFooterWidth = document.getElementById('tdFooter').offsetWidth + "px";
    divFooter.style.width = ffFooterWidth;
}

function displaySocialMedia()
{
    var  divSocialMedia = document.getElementById('divSocialMedia');
    if (isOnlineVersion == "true") 
    {
       divSocialMedia.style.display = "block";
    }
    else 
    {
        divSocialMedia.style.display = "none";
    }
}

function repositionRightPanel()
{
    var scrollRootDocElem = document.documentElement.scrollTop;
    var scrollRootBody = document.body.scrollTop;
    var scrollMax = null; parseInt(scrollRootBody);
    if (parseInt(scrollRootBody) < parseInt(scrollRootDocElem))
    { 
        scrollMax = parseInt(scrollRootDocElem); 
    }
    else { 
        scrollMax = parseInt(scrollRootBody); 
    }

    var elemCopyright = document.getElementById("banner");
    if (elemCopyright == null) { return; }
    var elem = document.getElementById("divPanel");
    var divPanelHeight = elem.clientHeight;
    if (!divPanelHeight){
        divPanelHeight = 792;
    }
    
    if (parseInt(scrollMax) > 89)
    { 
        elem.style.top = parseInt(scrollMax-15) + 1 + "px"; 
    }
    else 
    { 
        elem.style.top = 76 + "px"; 
    }
    
    if( elemCopyright.offsetTop < (scrollMax-15 + divPanelHeight)){
        elem.style.top = (elemCopyright.offsetTop - divPanelHeight) + "px"; 
    }
}

function repositionRightPanel2()
{
    var scrollRootDocElem = document.documentElement.scrollTop;
    var scrollRootBody = document.body.scrollTop;
    var scrollMax = null; parseInt(scrollRootBody);
    if (parseInt(scrollRootBody) < parseInt(scrollRootDocElem))
    { 
        scrollMax = parseInt(scrollRootDocElem); 
    }
    else { 
        scrollMax = parseInt(scrollRootBody); 
    }

    var elemCopyright = document.getElementById("divFooter");
    if (elemCopyright == null) { return; }
    var elem = document.getElementById("divPanel");
    var divPanelHeight = elem.clientHeight;
    if (!divPanelHeight){
        divPanelHeight = 792;
    }
    if (parseInt(scrollMax) > 89)
    { 
        elem.style.top = parseInt(scrollMax-89) + 1 + "px"; 
    }
    else 
    { 
        elem.style.top = 1 + "px"; 
    }

    if( elemCopyright.offsetTop < (scrollMax + divPanelHeight)){
        elem.style.top = (elemCopyright.offsetTop - divPanelHeight - 89 + 30) + "px"; 
    }
}