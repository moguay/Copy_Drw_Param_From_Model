var debug =  true;

var mGlob = pfcCreate("MpfcCOMGlobal");
var oSession = mGlob.GetProESession();
var server = oSession.GetActiveServer();

var fso1,fso2;
var f1,f2;

var arrayCache;

var spnFTR00205 = new Array();
spnFTR00205[0]= new Array("LENGTH",	"M2",	"M2.5",	"M3",				"M4",				"M5",				"M6",					"M8",				"M10");
spnFTR00205[1]= new Array("3",				"r",		"r",		"oos",				"oos",				"oos",				"oos",					"oos",				"oos");
spnFTR00205[2]= new Array("4",				"r",		"r",		"r",					"oos",				"oos",				"oos",					"oos",				"oos");
spnFTR00205[3]= new Array("5",				"r",		"r",		"r",					"r",					"oos",				"oos",					"oos",				"oos");
spnFTR00205[4]= new Array("6",				"r",		"r",		"HUA11454",		"NPR",				"r",					"oos",					"oos",				"oos");
spnFTR00205[5]= new Array("8",				"r",		"r",		"HUA11455",		"HUA11457",		"r",					"21628185",			"oos",				"oos");
spnFTR00205[6]= new Array("10",			"r",		"r",		"HUA11456",		"21628126",		"HUA22866",		"21628186",			"r",					"oos");
spnFTR00205[7]= new Array("12",			"r",		"r",		"HUA41153",		"HUA28825",		"HUA22802",		"21628187",			"HUA22842",		"r");
spnFTR00205[8]= new Array("16",			"r",		"r",		"HUA33183",		"HUA11458",		"HUA22867",		"21628188",			"HUA22843",		"21628248");
spnFTR00205[9]= new Array("20",			"r",		"r",		"HUA12569",		"NPR",				"HUA38120",		"21628189",			"HUA22844",		"21628249");
spnFTR00205[10]= new Array("25",			"oos",	"r",		"r",					"21628130",		"HUA22803",		"21628190",			"HUA22845",		"21628250");
spnFTR00205[11]= new Array("30",			"oos",	"oos",	"r",					"NPR",				"NPR",				"21628191",			"HUA22846",		"21628251");
spnFTR00205[12]= new Array("35",			"oos",	"oos",	"oos",				"NPR",				"HUA22868",		"21628192",			"HUA22847",		"21628252");
spnFTR00205[13]= new Array("40",			"oos",	"oos",	"oos",				"r",					"NPR",				"HUA36239",			"HUA22848",		"21628253");
spnFTR00205[14]= new Array("45",			"oos",	"oos",	"oos",				"oos",				"NPR",				"HUA27452",			"NPR",				"NPR");
spnFTR00205[15]= new Array("50",			"oos",	"oos",	"oos",				"oos",				"HUA22869",		"HUA27892",			"HUA22849",		"21628254");
spnFTR00205[16]= new Array("60",			"oos",	"oos",	"oos",				"oos",				"oos",				"21628195",			"HUA22850",		"NPR");

var prtFTR00205 = new Array();
prtFTR00205[0]= new Array("",		"",		"",		"",					"",					"",					"",						"",					"");
prtFTR00205[1]= new Array("",		"",		"",		"",					"",					"",					"",						"",					"");
prtFTR00205[2]= new Array("",		"",		"",		"",					"",					"",					"",						"",					"");
prtFTR00205[3]= new Array("",		"",		"",		"",					"",					"",					"",						"",					"");
prtFTR00205[4]= new Array("",		"",		"",		"hua1145401",	"",					"",					"",						"",					"");
prtFTR00205[5]= new Array("",		"",		"",		"hua1145501",	"hua1145701",	"",					"2162818501",		"",					"");
prtFTR00205[6]= new Array("",		"",		"",		"hua1145601",	"2162812601",	"hua2286601",	"2162818601",		"",					"");
prtFTR00205[7]= new Array("",		"",		"",		"nve1408500",	"hua2882501",	"hua2280201",	"2162818701",		"hua2284201",	"");
prtFTR00205[8]= new Array("",		"",		"",		"hua3318301",	"hua1145801",	"hua2286701",	"2162818801",		"hua2284301",	"2162824801");
prtFTR00205[9]= new Array("",		"",		"",		"hua1256901",	"",					"hua3812001",	"2162818901",		"hua2284401",	"2162824901");
prtFTR00205[10]= new Array("",		"",		"",		"",					"2162813001",	"hua2280301",	"2162819001",		"hua2284501",	"2162825001");
prtFTR00205[11]= new Array("",		"",		"",		"",					"",					"",					"2162819101",		"hua2284601",	"2162825101");
prtFTR00205[12]= new Array("",		"",		"",		"",					"",					"hua2286801",	"2162819201",		"hua2284701",	"2162825201");
prtFTR00205[13]= new Array("",		"",		"",		"",					"",					"",					"hua3623901",		"hua2284801",	"2162825301");
prtFTR00205[14]= new Array("",		"",		"",		"",					"",					"",					"hua2745201",		"",					"");
prtFTR00205[15]= new Array("",		"",		"",		"",					"",					"hua2286901",	"hua2789201",		"hua2284901",	"2162825401");
prtFTR00205[16]= new Array("",		"",		"",		"",					"",					"",					"2162819501",		"hua2285001",		"");
function Main(){
    try {
        // Browser window default size
        var browserSize = oSession.CurrentWindow.GetBrowserSize();
        if (browserSize > 35.0) {
            browserSize = 35.0;
            oSession.CurrentWindow.SetBrowserSize(0.0);
            oSession.CurrentWindow.SetBrowserSize(browserSize);
        }
    } catch(e) {}
    
    InitSaveList();

	var strOutputTable = "<table id='ftr00205'><thead><tr>";
	for (var j=0; j < spnFTR00205[0].length; j++){
		strOutputTable += "<th>" + spnFTR00205[0][j] + "</th>";
		// printDebug(j);
	}
	strOutputTable += "</tr></thead>";
	strOutputTable+= "<tbody>";
	// alert(strOutputTable);
	for (var i = 1; i < spnFTR00205.length; i++) {
		var strScrew = "<tr class='" + ( i % 2  > 0 ? "odd" : "even") + "'>";
		for (var j = 0;  j < spnFTR00205[i].length; j++) {
			// debugPrint (i + "|" + j);
			switch (spnFTR00205[i][j]) {
				case "r":
					// printDebug(i + "/" + j  + ": " +  spnFTR00205[i][j] + "| restricted");
					strScrew += "<td class='restricted'>&nbsp;</td>";
                    AddSaveList (spnFTR00205[i][j],"",(i*(spnFTR00205[i].length-1))-spnFTR00205[i].length+j,"","","");
					break;
				case "oos":
					// printDebug(i + "/" + j  + ": " +  spnFTR00205[i][j] + "| out of scope");
					strScrew += "<td class='outOfScope'>&nbsp;</td>";
                    AddSaveList (spnFTR00205[i][j],"",(i*(spnFTR00205[i].length-1))-spnFTR00205[i].length+j,"","","");
					break;
				case "NPR":
					// printDebug(i + "/" + j  + ": " +  spnFTR00205[i][j] + "| NPR");
					strScrew += "<td>NPR</td>";
                    AddSaveList (spnFTR00205[i][j],"",(i*(spnFTR00205[i].length-1))-spnFTR00205[i].length+j,"","","");
					break;
				default:
					// printDebug(i + "/" + j  + ": " +  spnFTR00205[i][j] + "|" + spnFTR00205[i][j] );
					// printDebug (i);
					if (j == 0 ){
						strScrew += "<td>" + spnFTR00205[i][j] + "</td>";
					} else {
                        //alert(data[(i*(spnFTR00205[i].length-1))-spnFTR00205[i].length+j + 1] + ((i*(spnFTR00205[i].length-1))-spnFTR00205[i].length+j));
                        arrayCache = data[(i*(spnFTR00205[i].length-1))-spnFTR00205[i].length+j + 1].toString().split("|");
                        
                        printDebug (arrayCache[0]);
                        //alert(arrayCache);
                        //wtPub = checkWS(prtFTR00205[i][j]);
                        
                        if (arrayCache[4].indexOf("wtpub") != -1) { wtPub = arrayCache[4]; }
                        else { wtPub = checkWS(prtFTR00205[i][j]); }
						strScrew += "<td><a class='draggable' se_part_number='"+spnFTR00205[i][j]+"' title='Description' href='"+wtPub+"'>" + spnFTR00205[i][j] + "</a></br><a href=''>&lt;info&gt;</a></td>";
                        AddSaveList (spnFTR00205[i][j],prtFTR00205[i][j]+'.prt',(i*(spnFTR00205[i].length-1))-spnFTR00205[i].length+j,"Description",wtPub,"INFO");
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

function InitSaveList () {

// HUA42020|nve4938500.prt|null|SCREW CS HXL SK M6 X 20 STL|wtpub://MDM Production/Libraries/Rondo_Library/Mechanical Components/Fasteners/Screw/General Screw Metric/nve4938500.prt|http://pfrxiri10.fr.schneider-electric.com/MDM/servlet/WindchillAuthGW/wt.enterprise.URLProcessor/URLTemplateAction?action=ObjProps&oid=OR%3Awt.epm.EPMDocument%3A2305292731&u8=1
// V12130015|v1213001501.prt|null|MSC HXG HD ISO4017 M12X45 8.8 Zn|wtpub://MDM Production/Libraries/Rondo_Library/Mechanical Components/Fasteners/Screw/General Screw Metric/v1213001501.prt|http://pfrxiri10.fr.schneider-electric.com/MDM/servlet/WindchillAuthGW/wt.enterprise.URLProcessor/URLTemplateAction?action=ObjProps&oid=OR%3Awt.epm.EPMDocument%3A1848552405&u8=1

    try {
        fso1 = new ActiveXObject("Scripting.FileSystemObject");
        f1 = fso1.OpenTextFile("FTR00205List.lst", 1);
        cache = f1.ReadAll();
        f1.Close();
        
    
        data = $.csv.toArrays(cache);
        //printDebug(data[0]);
        
        f1 = fso1.GetFile("FTR00205List.lst");
        f1.Delete();
    } catch(e) {
        alert(e);
    }
    
    try {
        fso2 = new ActiveXObject("Scripting.FileSystemObject");
        f2 = fso2.CreateTextFile("FTR00205List.lst", true);
        f2.WriteLine('Schneider Item No.;partNumber;description;file from delete row');
    }  catch(e) {
        alert(e);
    }
            
}

function AddSaveList (SPN,PART,CHECKED,DESC,WTPUB,INFO) {
    //alert(SPN + '|' + PART + '|' + CHECKED + '|' + DESC + '|' + WTPUB + '|' + INFO);
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