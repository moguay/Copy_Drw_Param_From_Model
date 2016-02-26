var mGlob = pfcCreate("MpfcCOMGlobal");
var objSession = mGlob.GetProESession();
var objCurrentModel = objSession.CurrentModel;
var arrModelHoles = objCurrentModel.ListFeaturesByType (false, pfcCreate ("pfcFeatureType").FEATTYPE_HOLE);
var objCurrentWindow = objSession.CurrentWindow;

function Main() {
	var featStatusActive = pfcCreate("pfcFeatureStatus").FEAT_ACTIVE;
	var UI = document.getElementById("UI");
	var holesTable = document.getElementById("holesTable");
	
	objCurrentModel.RetrieveView("Default");
	if (objCurrentModel.Descr.Type == pfcCreate("pfcModelType").MDL_PART & objCurrentModel.GetParam("SE_MATERIAL_1").Value.StringValue.substring(0, 8) == "HUA17334" ) {
		UI.innerHTML = objCurrentModel.InstanceName + " is a PCB<br />Material " + objCurrentModel.GetParam("SE_MATERIAL_1").Value.StringValue + "<br />" + arrModelHoles.Count + " holes found in the model to process<br />";
		var outputTable = "<table><thead><tr><th>ID</th><th>State</th><th>Ecad hole type</th></tr></thead><tbody>";
		for (var i = 0; i < arrModelHoles.Count; i++) {
			if (arrModelHoles.Item(i).Status == featStatusActive){
					if (arrModelHoles.Item(i).GetParam("ECAD_HOLE_TYPE") == void null) {
					var holeState = "New Hole";
					arrModelHoles.Item(i).CreateParam("ECAD_HOLE_TYPE", createParamValueFromString("NPTH"));
				} else {
						if (arrModelHoles.Item(i).GetParam("ECAD_HOLE_TYPE").Value.StringValue == "") {
							var holeState = "No value";
						} else {
							var holeState = "";
						} 
				}
				var holeType = arrModelHoles.Item(i).GetParam("ECAD_HOLE_TYPE").Value.StringValue;
				outputTable += "<tr class='" + ( i % 2  > 0 ? "odd" : "even") + "'>";
				outputTable += "<td ><button onmouseover='highlightHole (" + i + ")' >" + arrModelHoles.Item(i).Id + "</button></td>";
				outputTable += "<td>" + holeState + "</td>";
				outputTable += "<td><select id='" + i + "' onchange='updateValue(this)' ><option value='NPTH' " + (holeType == "NPTH" ? "selected" : "") + ">NPTH</option><option value='PTH' " + (holeType == "PTH" ? "selected" : "") + ">PTH</option></select>";
				outputTable += "</tr>";
				// outputTable += "<tr class='" + ( i % 2  > 0 ? "odd" : "even") + "'><td class='theID'>" + arrModelHoles.Item(i).Id + "</td><td>" + holeState + "</td><td><select id='" + i + "' onchange='updateValue(this)' ><option value='NPTH' " + (holeType == "NPTH" ? "selected" : "") + ">NPTH</option><option value='PTH' " + (holeType == "PTH" ? "selected" : "") + ">PTH</option></select>";
			}
		}
		holesTable.innerHTML = outputTable + "</tbody></table>";
		addOnMouseOver();
	} else {
		UI.innerHTML = "The model  you are trying to process is not a PCB.";
	}		
}

function highlightHole (holeID) {
	objCurrentWindow.Repaint();
	pfcCreate ("MpfcSelect").CreateModelItemSelection(objCurrentModel.GetItemById(pfcCreate("pfcModelItemType").ITEM_FEATURE, arrModelHoles.Item(holeID).Id), null).Highlight(pfcCreate ("pfcStdColor").COLOR_PRESEL_HIGHLIGHT);
}	

function updateValue (element) {
	arrModelHoles.Item(parseInt(element.id)).GetParam("ECAD_HOLE_TYPE").Value = pfcCreate("MpfcModelItem").CreateStringParamValue(element.value);
	objSession.GetModelWindow(Model).Refresh(); 
 }


function pfcIsWindows () {
	if (navigator.appName.indexOf ("Microsoft") != -1) {
		return true;
	} else {
		return false;
	}
}
function pfcCreate (className)  {
	if (!pfcIsWindows()) {
		netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
	}
	if (pfcIsWindows()) {
		return new ActiveXObject ("pfc."+className);
	} else {
		ret = Components.classes ["@ptc.com/pfc/" + className + ";1"].createInstance();
		return ret;
	}
}

function createParamValueFromString (s /* string */){
if (s.indexOf (".") == -1) {
          var i = parseInt (s);
          if (!isNaN(i))
        return pfcCreate ("MpfcModelItem").CreateIntParamValue(i);
        }
      else
        {
          var d = parseFloat (s);
          if (!isNaN(d))
        return pfcCreate ("MpfcModelItem").CreateDoubleParamValue(d);
        }
      if (s.toUpperCase() == "Y" || s.toUpperCase ()== "TRUE")
        return pfcCreate ("MpfcModelItem").CreateBoolParamValue(true);

      if (s.toUpperCase() == "N" || s.toUpperCase ()== "FALSE")
        return pfcCreate ("MpfcModelItem").CreateBoolParamValue(false);

      return pfcCreate ("MpfcModelItem").CreateStringParamValue(s);
    }

// function addOnMouseOver (){
	// var tableCells = document.getElementsByTagName('td');
	// for ( i in tableCells) {
		// if (tableCells[i].className == "theID") {
			// alert (">" + tableCells[i].innerHTML + "<");
			// var theIDValue = parseInt(tableCells[i].innerHTML);
			// tableCells[i].attachEvent("onmouseover",highlightHole(theIDValue));
		// }
	// }
// }
