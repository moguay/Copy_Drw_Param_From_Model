var mGlob = pfcCreate("MpfcCOMGlobal");
var oSession = mGlob.GetProESession();
var Model = oSession.CurrentModel;
var modelHoles = Model.ListFeaturesByType (false, pfcCreate ("pfcFeatureType").FEATTYPE_HOLE);

function Main() {
	//Initialization Commands
	var UI = document.getElementById("UI");
	var holesTable = document.getElementById("holesTable");
	if (Model.Descr.Type == pfcCreate("pfcModelType").MDL_PART & Model.GetParam("SE_MATERIAL_1").Value.StringValue.substring(0, 8) == "HUA17334" ) {
		UI.innerHTML = Model.InstanceName + " is a PCB<br />Material " + Model.GetParam("SE_MATERIAL_1").Value.StringValue + "<br />" + modelHoles.Count + " holes found in the model to process<br />";
		var outputTable = "<table><thead><tr><th >ID</th><th>State</th><th>Ecad hole type</th></tr></thead><tbody>";
		for (var i = 0; i < modelHoles.Count; i++) {
				if (modelHoles.Item(i).GetParam("ECAD_HOLE_TYPE") == void null) {
					var holeState = "Not set";
					var holeValue= "NPTH";
					modelHoles.Item(i).CreateParam("ECAD_HOLE_TYPE", createParamValueFromString(holeValue));
				} else {
					var holeState = "";
					var holeValue= modelHoles.Item(i).GetParam("ECAD_HOLE_TYPE").Value.StringValue;
				}
				outputTable += "<tr class='" + ( i % 2  > 0 ? "odd" : "even") + "'><td>" + modelHoles.Item(i).Id + "</td><td>" + holeState + "</td><td><input id='" + i + "' type='text'  value='" + holeValue + "' onfocus='highlightHole(this.id)' onchange='updateHoleValue(this.id, this.value)'></td></tr>";
		}
		holesTable.innerHTML = outputTable + "</tbody></table>";
		oSession.GetModelWindow(Model).Repaint(); 
	} else {
		UI.innerHTML = "The model  you are trying to process is not a PCB.";
	}
}

function highlightHole (holeID) {
	oSession.GetModelWindow(Model).Repaint(); 
	pfcCreate ("MpfcSelect").CreateModelItemSelection(Model.GetItemById(pfcCreate("pfcModelItemType").ITEM_FEATURE, modelHoles.Item(parseInt(holeID)).Id), null).Display();
}

function updateHoleValue (holeID, newValue) {
	alert (holeID + ' ' + newValue);
	modelHoles.Item(holeID).GetParam("ECAD_HOLE_TYPE").Value = pfcCreate("MpfcModelItem").CreateStringParamValue(newValue);
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

function CreateParamValue(typeAsString,val)
    {
        switch (typeAsString)
        {
            case "STRING":
            return pfcCreate("MpfcModelItem").CreateStringParamValue(val);
            break;
            case "NUMBER":
            return pfcCreate("MpfcModelItem").CreateDoubleParamValue(val);
            break;
            case "BOOLEAN":
            return pfcCreate("MpfcModelItem").CreateBoolParamValue(val);
            break;
            case "INTEGER":
            return pfcCreate("MpfcModelItem").CreateIntParamValue(val);
            break;
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

	
