var debug =  true;

var mGlob = pfcCreate("MpfcCOMGlobal");
var oSession = mGlob.GetProESession();
var prtInWS = [];
var server ;
var fileListType = pfcCreate("pfcFileListOpt").FILE_LIST_ALL;

try {
	var server = oSession.GetActiveServer();
	var activeWS = "wtws://" + server.Alias + "/" + server.ActiveWorkspace + "/";
	// alert(activeWS);
}
catch(e){
	printDebug("no server");
}
function Main(strFTR){
	try {
		// Browser window default size
		var browserSize = oSession.CurrentWindow.GetBrowserSize();
		if (browserSize != 70.0) {
			browserSize = 70.0;
			oSession.CurrentWindow.SetBrowserSize(0.0);
			oSession.CurrentWindow.SetBrowserSize(browserSize);
		}
	}
	catch(e) {
	}
	for (i = 0, j = spnPRT.length; i < j; i++){
		if (spnPRT[i][0] == strFTR) {
			prtInWS[spnPRT[i][1]] = spnPRT[i][2];
			// alert(strFTR + "|" + spnPRT[i][2]);
		}
	}	
	// alert('');
	var filesList = oSession.ListFiles("*.prt", fileListType, activeWS);
	for (var k in prtInWS){
		// printDebug(prtInWS[k]);
		var tempPRT = activeWS + prtInWS[k].substring(prtInWS[k].lastIndexOf("/") + 1);
		// alert(tempPRT);
		// alert(filesList.Item(0));
		for (i = 0, j = filesList.Count; i < j; i++) {
			if (filesList.Item(i) == tempPRT){
				prtInWS[k] = tempPRT;
			}
		}
	}
	var strOutputTable = "<table id='ftr'><thead><tr>";
	for (j = 0; j < spnFTR[0].length; j++){
		strOutputTable += "<th>" + spnFTR[0][j] + "</th>";
		// printDebug(j);
	}
	strOutputTable += "</tr></thead>";
	strOutputTable+= "<tbody>";
	// alert(strOutputTable);
	for ( i = 1; i < spnFTR.length; i++) {
		var strScrew = "<tr class='" + ( i % 2  > 0 ? "odd" : "even") + "'>";
		for ( j = 0;  j < spnFTR[i].length; j++) {
			// debugPrint (i + "|" + j);
			switch (spnFTR[i][j]) {
				case "r":
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "| restricted");
					strScrew += "<td class='restricted'>&nbsp;</td>";
					break;
				case "oos":
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "| out of scope");
					strScrew += "<td class='outOfScope'>&nbsp;</td>";
					break;
				case "NPR":
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "| NPR");
					strScrew += "<td>NPR</td>";
					break;
				default:
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "|" + spnFTR[i][j] );
					// printDebug (i);
					if (j == 0 ){
						strScrew += "<td>" + spnFTR[i][j] + "</td>";
					} else {
						// strScrew += "<td><a class='draggable' se_part_number='"+spnFTR[i][j]+"' title='Description' href='"+wtPub+"'>" + spnFTR[i][j] + "</a></br><a href=''>&lt;info&gt;</a></td>";
						// printDebug (prtInWS["ST411-034-160"]);
						strScrew += "<td><a class='draggable' se_part_number='"+spnFTR[i][j]+"' title='"+spnFTR[i][j]+"' href='"+ prtInWS[spnFTR[i][j]]+"'>" + spnFTR[i][j] + "</a></td>";
					}
			}
		}
	strOutputTable += strScrew +"</tr>";
	}
	strOutputTable+= "</tbody></table>";
	document.getElementById('holesTable').innerHTML = strOutputTable;

    SetDragDropEvents();

}

function printDebug(strOutput){
	if (debug){
		strOutput = strOutput.replace("<", "&lt;");
		strOutput = strOutput.replace(">", "&gt;");
		document.getElementById("debug").innerHTML += "<p>" + strOutput + "</p>";
	}
}
function pfcIsWindows () {
	if (navigator.appName.indexOf ("Microsoft") != -1) {
		return true;
	} else {
		return false;
	}
}
// Optimized pfcCreate
function pfcCreate (className) {
	if (!pfcIsWindows())
		netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
	if (className.match(/^M?pfc/)) {
		// Check global object cache first, then return that object
		//
		try {
			if (className in global_class_cache) {
				return global_class_cache[className];
			}
		}
		catch (e) {
			// Probably no global_class_cache yet
			global_class_cache = new Object();
		}
	}
	// Not in global object cache, create object
	var obj = null;
	if (pfcIsWindows()) {
		obj = new ActiveXObject("pfc."+className);
	} else {
		obj = Components.classes ["@ptc.com/pfc/" + className + ";1"].createInstance();
	}
	// Return created object
	//
	if (className.match(/^M?pfc/)) {
		global_class_cache[className] = obj;
	}
	return obj;
}

//// LIB SPN

function isProEEmbeddedBrowser ()
{
	if (top.external && top.external.ptc)
		return true;
	else
		return false;
}

function pfcGetProESession ()
{
	if (!isProEEmbeddedBrowser ())
	{
		throw new Error ("Not in embedded browser.  Aborting...");
	}

	// Security code
	if (!pfcIsWindows())
		netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");

	var glob = pfcCreate ("MpfcCOMGlobal");
	return glob.GetProESession();
}

function pfcGetExceptionType (err)
{
	if (pfcIsWindows())
		return (err.description);
	else
	{
		errString = err.message;
		// This should remove the XPCOM prefix ("XPCR_C")
		if (errString.search ("XPCR_C") == 0)
			return (errString.substr(6));
		else 
			return (errString);
	}
}


/***********************************************************************************************
 * If current model is an assembly, create a new parameter called LATEST_PART_NUMBER
 * to temporary store related SE_PART_NUMBER of dragged URL
 **********************************************************************************************/
function handleDragStart (e) {
	if(!e) {
		var e = window.event;
	}
	if(!e.target){
		e.target = e.srcElement;
	}

	latest_se_part_number = e.target.se_part_number

	if (!pfcIsWindows())
		netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");

	session = pfcGetProESession ();
	mdl = session.getActiveModel();
	if (mdl) {
		mdl_type = mdl.Type;
		assy_type = pfcCreate("pfcModelType").MDL_ASSEMBLY

		if (mdl_type == assy_type) {
			param1 = pfcCreate ("MpfcModelItem").CreateStringParamValue(latest_se_part_number);
			p = mdl.GetParam("LATEST_PART_NUMBER");
			if (p) {
				// Update parameter
				p.Value = param1;
			} else {
				// Create parameter
				mdl.CreateParam("LATEST_PART_NUMBER", param1 );
			}
		}
	}
}

/***********************************************************************************************
 * Search all "draggable" objects
 * Associate a startdrag event to each
 **********************************************************************************************/
function SetDragDropEvents () {
	strout = ""

	if (document.getElementsByTagName) {
		elems = document.getElementsByTagName('*')

		cnt = 0
		for (i in elems) {
			if (elems[i] && elems[i].className) {
				strout = strout  + "\n" + elems[i].className
				if (elems[i].className == "draggable") {
					cnt = cnt+1
					// Register events
					draggable_url = elems[i]
					
					// Tricky code to handle IE8 proprietary event mechanism
					if (draggable_url.addEventListener) {
						// TO DO
						// toto.addEventListener('dragstart', handleDragStart, false);
					} else {
						draggable_url.attachEvent("ondragstart", function f(draggable_url) {handleDragStart(draggable_url)});
					}
				}
			}
		}
	}
}