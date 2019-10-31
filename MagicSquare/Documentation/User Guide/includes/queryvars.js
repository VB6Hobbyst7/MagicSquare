/*
==============================================================================
Iuga Alexandru - Sept 2005
------------------------------------------------------------------------------

A class to parse the QueryString af the page.

Fields:
	aQueryVars - a 2D table (2 x QVCount) containing the pairs name-value of the
					QueryString variables.
Methods:
	GetValue - returns the value of a QV.
==============================================================================
*/
function QueryVars(strQueryString)
{
	if (strQueryString == undefined) strQueryString = document.location.href.split("?")[1];
	
	this.aQueryVars = new Array();
	
	if (strQueryString != undefined)
	{
		var aTmp;
		
		this.aQueryVars = strQueryString.split("&");
		for (i=0; i<this.aQueryVars.length; i++)
		{
			aTmp = this.aQueryVars[i].split("=");
			this.aQueryVars[i] = new Array(aTmp[0], aTmp[1]);
		}
	}
	
	this.GetValue = function(strVarName)
	{
		for (i=0; i<this.aQueryVars.length; i++)
		{
			if (this.aQueryVars[i][0] == strVarName) return this.aQueryVars[i][1];
		}
		return undefined;
	}
}