var debug =  true;

var mGlob = pfcCreate("MpfcCOMGlobal");
var oSession = mGlob.GetProESession();
var server = oSession.GetActiveServer();

var fso1,fso2;
var f1,f2;

var arrayCache;

function Main(){
	try {
		openDB();
        // Browser window default size
        var browserSize = oSession.CurrentWindow.GetBrowserSize();
        if (browserSize > 35.0) {
            browserSize = 35.0;
            oSession.CurrentWindow.SetBrowserSize(0.0);
            oSession.CurrentWindow.SetBrowserSize(browserSize);
        }
    } catch(e) {}
    
    InitSaveList();

	var strOutputTable = "<table id='ftr'><thead><tr>";
	for (var j=0; j < spnFTR[0].length; j++){
		strOutputTable += "<th>" + spnFTR[0][j] + "</th>";
		// printDebug(j);
	}
	strOutputTable += "</tr></thead>";
	strOutputTable+= "<tbody>";
	// alert(strOutputTable);
	for (var i = 1; i < spnFTR.length; i++) {
		var strScrew = "<tr class='" + ( i % 2  > 0 ? "odd" : "even") + "'>";
		for (var j = 0;  j < spnFTR[i].length; j++) {
			// debugPrint (i + "|" + j);
			switch (spnFTR[i][j]) {
				case "r":
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "| restricted");
					strScrew += "<td class='restricted'>&nbsp;</td>";
                    AddSaveList (spnFTR[i][j],"",(i*(spnFTR[i].length-1))-spnFTR[i].length+j,"","","");
					break;
				case "oos":
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "| out of scope");
					strScrew += "<td class='outOfScope'>&nbsp;</td>";
                    AddSaveList (spnFTR[i][j],"",(i*(spnFTR[i].length-1))-spnFTR[i].length+j,"","","");
					break;
				case "NPR":
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "| NPR");
					strScrew += "<td>NPR</td>";
                    AddSaveList (spnFTR[i][j],"",(i*(spnFTR[i].length-1))-spnFTR[i].length+j,"","","");
					break;
				default:
					// printDebug(i + "/" + j  + ": " +  spnFTR[i][j] + "|" + spnFTR[i][j] );
					// printDebug (i);
					if (j == 0 ){
						strScrew += "<td>" + spnFTR[i][j] + "</td>";
					} else {
                        //alert(data[(i*(spnFTR[i].length-1))-spnFTR[i].length+j + 1] + ((i*(spnFTR[i].length-1))-spnFTR[i].length+j));
                        try {
                            arrayCache = data[(i*(spnFTR[i].length-1))-spnFTR[i].length+j + 1].toString().split("|");
                            if (arrayCache[4].indexOf("wtpub") != -1) { wtPub = arrayCache[4]; }
                            else { wtPub = checkWS(prtFTR[i][j]); }
                        } catch(e) {
                            wtPub = checkWS(prtFTR[i][j]); 
                        }
                        
                        //printDebug (arrayCache[0]);
                        //alert(arrayCache);
                        //wtPub = checkWS(prtFTR[i][j]);
                        
						// strScrew += "<td><!--<a class='draggable' se_part_number='"+spnFTR[i][j]+"' title='Description' href='"+wtPub+"'>-->" + spnFTR[i][j] + "<!--</a></br><a href=''>&lt;info&gt;</a>--></td>";
						strScrew += "<td><a class='draggable' se_part_number='"+spnFTR[i][j]+"' title='Description' href='"+wtPub+"'>" + spnFTR[i][j] + "</a></br><a href=''>&lt;info&gt;</a></td>";



                        AddSaveList (spnFTR[i][j],prtFTR[i][j]+'.prt',(i*(spnFTR[i].length-1))-spnFTR[i].length+j,"Description",wtPub,"INFO");
					}
			}
		}
	strOutputTable += strScrew +"</tr>";
	}
	strOutputTable+= "</tbody></table>";
	document.getElementById('holesTable').innerHTML += strOutputTable;

    SetDragDropEvents();

    CloseSaveList();
}

function checkWS(strPRTName) {
    var wslink;

    var aryPRTName = pfcCreate ("stringseq");
    aryPRTName.set (0, strPRTName.toLowerCase() + ".prt");
    aryPRTName.set (1, strPRTName.toUpperCase() + ".PRT");

    for (i=0;i<2 && wslink==void null;i++) {
        try {
            //printDebug (aryPRTName.item(i));
            wslink = server.GetAliasedUrl(aryPRTName.item(i));
            // wslink = 'disabled';
            //printDebug (wslink);
        }
        catch (e) { }
    }

    if (wslink == void null) {
        printDebug ("wsLink not found: " + strPRTName);
        return;
    }
    return wslink;
}
function printDebug(strOutput){
	if (debug){
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
function openDB (){
	try {
		fso1 = new ActiveXObject("Scripting.FileSystemObject");
        var scripts= document.getElementsByTagName('script');
		var path= scripts[scripts.length-1].src.split('?')[0];      // remove any ?query
		var mydir= path.split('/').slice(0, -1).join('/')+'/';
		f1 = fso1.OpenTextFile("./standardized.list.txt", 1);
        var list = f1.ReadAll();
		printDebug(list);
	}
	catch(e){
		printDebug(e);
	}
}

function InitSaveList () {

// HUA42020|nve4938500.prt|null|SCREW CS HXL SK M6 X 20 STL|wtpub://MDM Production/Libraries/Rondo_Library/Mechanical Components/Fasteners/Screw/General Screw Metric/nve4938500.prt|http://pfrxiri10.fr.schneider-electric.com/MDM/servlet/WindchillAuthGW/wt.enterprise.URLProcessor/URLTemplateAction?action=ObjProps&oid=OR%3Awt.epm.EPMDocument%3A2305292731&u8=1
// V12130015|v1213001501.prt|null|MSC HXG HD ISO4017 M12X45 8.8 Zn|wtpub://MDM Production/Libraries/Rondo_Library/Mechanical Components/Fasteners/Screw/General Screw Metric/v1213001501.prt|http://pfrxiri10.fr.schneider-electric.com/MDM/servlet/WindchillAuthGW/wt.enterprise.URLProcessor/URLTemplateAction?action=ObjProps&oid=OR%3Awt.epm.EPMDocument%3A1848552405&u8=1

    try {
        fso1 = new ActiveXObject("Scripting.FileSystemObject");
        f1 = fso1.OpenTextFile(document.title + ".lst", 1);
        cache = f1.ReadAll();
        f1.Close();
        
    
        data = $.csv.toArrays(cache);
        //printDebug(data[0]);
        
        f1 = fso1.GetFile(document.title + ".lst");
        f1.Delete();
    } catch(e) {
        // alert(e);
    }
    
    try {
        fso2 = new ActiveXObject("Scripting.FileSystemObject");
        f2 = fso2.CreateTextFile(document.title + ".lst", true);
        f2.WriteLine('Schneider Item No.;partNumber;description;file from delete row');
    }  catch(e) {
        // alert(e);
    }
            
}

function AddSaveList (SPN,PART,CHECKED,DESC,WTPUB,INFO) {
    // printDebug(SPN + '|' + PART + '|' + CHECKED + '|' + DESC + '|' + WTPUB + '|' + INFO);
    try {
        f2.WriteLine (SPN + '|' + PART + '|' + CHECKED + '|' + DESC + '|' + WTPUB + '|' + INFO);
    } catch(e) {
        printDebug (e);
    }
    
}

function CloseSaveList () {
    try {
        f2.Close();
    } catch(e) {
        printDebug (e);
    }
}