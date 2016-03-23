var debug =  true;

var mGlob = pfcCreate("MpfcCOMGlobal");
var oSession = mGlob.GetProESession();
var server = oSession.GetActiveServer();

function Main(){
	var filesList = osession.ListFiles("*.prt", pfcCreate("pfcFileListOpt").FILE_LIST_LATEST, "wtws://" + server.Alias + "/" + server.ActiveWorkspace + "/" );
	alert (filesList.Count);
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

