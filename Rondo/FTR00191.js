var debug =  true;

var mGlob = pfcCreate("MpfcCOMGlobal");
var oSession = mGlob.GetProESession();
var spnFTR00191 = new Array();

spnFTR00191[0]= new Array("LENGTH",	"M1.6",	"M2",	"M2.5",	"M3",				"M4",				"M5",				"M6",					"M8",				"M10",				"M12", 				"M16", 				"M20",				"M24");
spnFTR00191[1]= new Array("2",			"r",		"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[1]= new Array("3",			"r",		"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[2]= new Array("4",			"r",		"r",		"oos",	"oos",				"oos",				"oos",				"oos",					"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[3]= new Array("5",			"r",		"r",		"r",		"oos",				"oos",				"oos",				"oos",					"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[4]= new Array("6",			"r",		"r",		"r",		"HUA11440",		"oos",				"oos",				"oos",					"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[5]= new Array("8",			"r",		"r",		"r",		"21431095",		"HUA11442",		"oos",				"oos",					"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[6]= new Array("10",			"r",		"r",		"r",		"HUA11441",		"HUA11486",		"HUA11447",		"oos",					"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[7]= new Array("12",			"r",		"r",		"r",		"HUA12593",		"HUA14556",		"HUA11448",		"HUA12367",			"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[8]= new Array("16",			"r",		"r",		"r",		"HUA14552",		"HUA11443",		"HUA14562",		"HUA12371",			"oos"				"oos",				"oos",				"oos",				"oos",				"oos");
spnFTR00191[9]= new Array("20",			"oos",	"r",		"r",		"21431099",		"HUA11444",		"HUA13753",		"HUA12368",			"HUA12372",		"HUA13745",		"oos",				"oos",				"oos",				"oos");
spnFTR00191[10]= new Array("25",		"oos",	"oos",	"r",		"HUA14554",		"HUA14557",		"HUA14563",		"HUA14570",			"HUA12373",		"HUA13746",		"HUA12376",		"oos",				"oos",				"oos");
spnFTR00191[11]= new Array("30",		"oos",	"oos",	"oos",	"21431101",		"HUA11445",		"HUA14564",		"HUA14571",			"HUA14578",		"HUA12374",		"HUA14609",		"HUA14630",		"oos",				"oos");
spnFTR00191[12]= new Array("35",		"oos",	"oos",	"oos",	"oos",				"HUA14558",		"HUA14565",		"HUA12369",			"HUA14579",		"HUA13747",		"HUA14610",		"HUA14631",		"oos",				"oos");
spnFTR00191[13]= new Array("40",		"oos",	"oos",	"oos",	"oos",				"HUA11446",		"HUA11449",		"HUA14572",			"HUA14591",		"HUA12375",		"HUA14611",		"HUA14632",		"r",					"oos");
spnFTR00191[14]= new Array("45",		"oos",	"oos",	"oos",	"oos",				"oos",				"HUA20515",		"V1109969",			"NPR",				"HUA13748",		"V12130015",	"NPR",				"r",					"oos");
spnFTR00191[15]= new Array("50",		"oos",	"oos",	"oos",	"oos",				"oos",				"HUA12594",		"HUA12370",			"HUA14592",		"HUA14600",		"HUA14612",		"HUA14633",		"r",					"r");
spnFTR00191[16]= new Array("55",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"HUA41333",			"NPR",				"21431639",		"NPR",				"NPR",				"r",					"r");
spnFTR00191[16]= new Array("60",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"HUA11451",			"HUA14593",		"HUA14601",		"HUA14613",		"HUA14634",		"r",					"r");
spnFTR00191[16]= new Array("65",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"NPR",				"NPR",				"NPR",				"NPR",				"r",					"r");
spnFTR00191[16]= new Array("70",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"HUA14594",		"HUA14602",		"HUA14614",		"NPR",				"r",					"r");
spnFTR00191[16]= new Array("80",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"HUA14595",		"HUA14604",		"HUA14615",		"HUA40537",		"r",					"r");
spnFTR00191[16]= new Array("90",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"HUA13750",		"HUA14616",		"r",					"r",					"r");
spnFTR00191[16]= new Array("100",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"HUA11453",		"HUA14617",		"r",					"r",					"r");
spnFTR00191[16]= new Array("110",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"r",					"r",					"r",					"r");
spnFTR00191[16]= new Array("120",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"r",					"r",					"r",					"r");
spnFTR00191[16]= new Array("130",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"oos",				"r",					"r",					"r");
spnFTR00191[16]= new Array("140",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"oos",				"r",					"r",					"r");
spnFTR00191[16]= new Array("150",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"oos",				"r",					"r",					"r");
spnFTR00191[16]= new Array("160",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"oos",				"r",					"r",					"r");
spnFTR00191[16]= new Array("180",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"oos",				"r",					"r",					"r");
spnFTR00191[16]= new Array("200",		"oos",	"oos",	"oos",	"oos",				"oos",				"oos",				"oos",					"oos",				"oos",				"oos",				"r",					"r",					"r");



var prtFTR00191 = new Array();
prtFTR00191[0]= new Array("LENGTH",	"M1.6",	"M2",	"M2.5",	"M3",				"M4",				"M5",				"M6",					"M8",				"M10",				"M12", 				"M16", 				"M20",				"M24");
prtFTR00191[1]= new Array("2",				"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[1]= new Array("3",				"",			"",			"",			"",						"",						"",						"",							""						"",						"",						"",						"",						"");
prtFTR00191[2]= new Array("4",				"",			"",			"",			"",						"",						"",						"",							""						"",						"",						"",						"",						"");
prtFTR00191[3]= new Array("5",				"",			"",			"",			"",						"",						"",						"",							""						"",						"",						"",						"",						"");
prtFTR00191[4]= new Array("6",				"",			"",			"",			"hua1144001",	"",						"",						"",							""						"",						"",						"",						"",						"");
prtFTR00191[5]= new Array("8",				"",			"",			"",			"2143109501",	"hua1144201",	"",						"",							""						"",						"",						"",						"",						"");
prtFTR00191[6]= new Array("10",			"",			"",			"",			"hua1144101",	"hua1148601",	"hua1144701",	"",							""						"",						"",						"",						"",						"");
prtFTR00191[7]= new Array("12",			"",			"",			"",			"hua1259301",	"hua1455601",	"hua1144801",	"hua1236701",		""						"",						"",						"",						"",						"");
prtFTR00191[8]= new Array("16",			"",			"",			"",			"hua1455201",	"hua1144301",	"HUA14562",		"hua1237101",		""						"",						"",						"",						"",						"");
prtFTR00191[9]= new Array("20",			"",			"",			"",			"2143109901",	"hua1144401",	"hua1375301",	"hua1236801",		"",						"HUA13745",		"",						"",						"",						"");
prtFTR00191[10]= new Array("25",			"",			"",			"",			"HUA14554",		"hua1455701",	"hua1456301",	"HUA14570",			"hua1237301",	"HUA13746",		"hua1237601",	"",						"",						"");
prtFTR00191[11]= new Array("30",			"",			"",			"",			"hua1455401",	"hua1144501",	"hua1456401",	"hua1457101",		"hua1457801",	"HUA12374",		"hua1460901",	"hua1463001",	"",						"");
prtFTR00191[12]= new Array("35",			"",			"",			"",			"",						"hua1455801",	"hua1456501",	"hua1236901",		"hua1457901",	"HUA13747",		"hua1461001",	"hua1463101",	"",						"");
prtFTR00191[13]= new Array("40",			"",			"",			"",			"",						"hua1144601",	"hua1144901",	"hua1457201",		"hua1459101",	"HUA12375",		"hua1461101",	"hua1463201",	"",						"");
prtFTR00191[14]= new Array("45",			"",			"",			"",			"",						"",						"hua2051501",	"v110996901",		"",						"HUA13748",		"v1213001501",	"",						"",						"");
prtFTR00191[15]= new Array("50",			"",			"",			"",			"",						"",						"hua1259401",	"hua1237001",		"hua1459201",	"HUA14600",		"hua1461201",	"hua1463301",	"",						"");
prtFTR00191[16]= new Array("55",			"",			"",			"",			"",						"",						"",						"nve2599900",		"",						"21431639",		"",						"",						"",						"");
prtFTR00191[16]= new Array("60",			"",			"",			"",			"",						"",						"",						"hua1145101",		"hua1459301",	"HUA14601",		"hua1461301",	"hua1463401",	"",						"");
prtFTR00191[16]= new Array("65",			"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("70",			"",			"",			"",			"",						"",						"",						"",							"hua1459401",	"HUA14602",		"hua1461401",	"",						"",						"");
prtFTR00191[16]= new Array("80",			"",			"",			"",			"",						"",						"",						"",							"hua1459501",	"HUA14604",		"hua1461501",	"nha8414400",	"",						"");
prtFTR00191[16]= new Array("90",			"",			"",			"",			"",						"",						"",						"",							"",						"HUA13750",		"hua1461601",	"",						"",						"");
prtFTR00191[16]= new Array("100",		"",			"",			"",			"",						"",						"",						"",							"",						"HUA11453",		"hua1461701",	"",						"",						"");
prtFTR00191[16]= new Array("110",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("120",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("130",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("140",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("150",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("160",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("180",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
prtFTR00191[16]= new Array("200",		"",			"",			"",			"",						"",						"",						"",							"",						"",						"",						"",						"",						"");
function Main(){
	var strOutputTable = "<table id='ftr00191'><thead><tr>";
	for (var j=0; j < spnFTR00191[0].length; j++){
		strOutputTable += "<th>" + spnFTR00191[0][j] + "</th>";
		// printDebug(j);
	}
	strOutputTable += "</tr></thead>";
	strOutputTable+= "<tbody>";
	// alert(strOutputTable);
	for (var i = 1; i < spnFTR00191.length; i++) {
		var strScrew = "<tr class='" + ( i % 2  > 0 ? "odd" : "even") + "'>";
		for (var j = 0;  j < spnFTR00191[i].length; j++) {
			// debugPrint (i + "|" + j);
			switch (spnFTR00191[i][j]) {
				case "r":
					// printDebug(i + "/" + j  + ": " +  spnFTR00191[i][j] + "| restricted");
					strScrew += "<td class='restricted'>&nbsp;</td>";
					break;
				case "oos":
					// printDebug(i + "/" + j  + ": " +  spnFTR00191[i][j] + "| out of scope");
					strScrew += "<td class='outOfScope'>&nbsp;</td>";
					break;
				case "NPR":
					// printDebug(i + "/" + j  + ": " +  spnFTR00191[i][j] + "| NPR");
					strScrew += "<td>NPR</td>";
					break;
				default:
					// printDebug(i + "/" + j  + ": " +  spnFTR00191[i][j] + "|" + spnFTR00191[i][j] );
					// printDebug (i);
					if (j == 0 ){
						strScrew += "<td>" + spnFTR00191[i][j] + "</td>";
					} else {
						checkWS(prtFTR00191[i][j]);
						strScrew += "<td><a href='wtpub://MDM Production/Libraries/Rondo_Library/Mechanical Components/Fasteners/Screw/General Screw Metric/" + prtFTR00191[i][j] + ".prt'>" + spnFTR00191[i][j] + "</a></td>";
					}
			}
		}
	strOutputTable += strScrew +"</tr>";
	}
	strOutputTable+= "</tbody></table>";
	document.getElementById('holesTable').innerHTML += strOutputTable;
}

function checkWS(strPRTName) {
	alert('test');
	// server = oSession.GetActiveServer();
	// var modelInWs = server.IsObjectCheckedOut(server.ActiveWorkspace,strPRTName + ".prt");
	// if (modelInWs != void null) {
		// prindDebug (strPRTName);
	// }
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
