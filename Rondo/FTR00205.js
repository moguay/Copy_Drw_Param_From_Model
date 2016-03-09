var debug =  true;

var mGlob = pfcCreate("MpfcCOMGlobal");
var oSession = mGlob.GetProESession();
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
					break;
				case "oos":
					// printDebug(i + "/" + j  + ": " +  spnFTR00205[i][j] + "| out of scope");
					strScrew += "<td class='outOfScope'>&nbsp;</td>";
					break;
				case "NPR":
					// printDebug(i + "/" + j  + ": " +  spnFTR00205[i][j] + "| NPR");
					strScrew += "<td>NPR</td>";
					break;
				default:
					// printDebug(i + "/" + j  + ": " +  spnFTR00205[i][j] + "|" + spnFTR00205[i][j] );
					// printDebug (i);
					if (j == 0 ){
						strScrew += "<td>" + spnFTR00205[i][j] + "</td>";
					} else {
						checkWS(prtFTR00205[i][j]);
						
						strScrew += "<td><a href='wtpub://MDM Production/Libraries/Rondo_Library/Mechanical Components/Fasteners/Screw/General Screw Metric/" + prtFTR00205[i][j] + ".prt'>" + spnFTR00205[i][j] + "</a></td>";
					}
			}
		}
	strOutputTable += strScrew +"</tr>";
	}
	strOutputTable+= "</tbody></table>";
	document.getElementById('holesTable').innerHTML += strOutputTable;
}

function checkWS(strPRTName) {
	server = oSession.GetActiveServer();
	
	try {
		printDebug(server.GetActiveWorkspace());
		var modelInWs = server.IsObjectCheckedOut(server.ActiveWorkspace,strPRTName + ".prt");
		if (modelInWs != void null) {
			printDebug ("in Ws: " + strPRTName);
		}
	}
	catch (e) {
		
		// printDebug (e);
		printDebug ("not in Ws: " + strPRTName);
		
	}
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
