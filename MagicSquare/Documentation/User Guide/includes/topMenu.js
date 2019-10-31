/*
(function(i, s, o, g, r, a, m) {
    i['GoogleAnalyticsObject'] = r; i[r] = i[r] || function() {
        (i[r].q = i[r].q || []).push(arguments)
    }, i[r].l = 1 * new Date(); a = s.createElement(o),
  m = s.getElementsByTagName(o)[0]; a.async = 1; a.src = g; m.parentNode.insertBefore(a, m)
})(window, document, 'script', 'https://www.google-analytics.com/analytics.js', 'ga');

ga('create', 'UA-22401154-1', 'auto');
ga('send', 'pageview');
*/

var active;
	function loadTopMenu(folderParentImages)
	{
		document.getElementById("tab1").src = folderParentImages + 'images/tab1_active.gif';	
		document.getElementById("tab2").src = folderParentImages + 'images/tab2_inactive.gif';	
		active = 1;
	}
	function changeActive(option, folderParentImages)
	{
				if (option == 1)
		{
			document.getElementById("tab1").src = folderParentImages + 'images/tab1_active.gif';	
			document.getElementById("tab2").src = folderParentImages + 'images/tab2_inactive.gif';
			parent.top.location = folderParentImages + "index.html";
			
		}
		else
		{
			document.getElementById("tab1").src = folderParentImages + 'images/tab1_inactive.gif';	
			document.getElementById("tab2").src = folderParentImages + 'images/tab2_active.gif';	
			parent.top.location = folderParentImages + "indexAPI.html";	
		}
		active = option;
	}
	function mouseover(option, folderParentImages){
		var objLink = document.getElementById("link" + option);
		if(objLink!=null){
			objLink.style.cursor = 'hand';
			objLink.style.cursor = 'pointer';
		}
		if(active != option){
			var objTab = document.getElementById("tab" + option);			
			if(objTab!=null)
				objTab.src = folderParentImages + 'images/tab' + option + '_active.gif';
		}			
	}
	function mouseout(option, folderParentImages){
		if(active != option){
			var objTab = document.getElementById("tab" + option);
			if(objTab!=null)
				objTab.src = folderParentImages + 'images/tab' + option + '_inactive.gif';
		}	
	}
