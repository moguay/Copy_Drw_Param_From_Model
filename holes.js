function Main() {
	//Initialization Commands
	var mGlob = pfcCreate("MpfcCOMGlobal");
	var oSession = mGlob.GetProESession();
	var Model = oSession.CurrentModel;
	var UI = document.getElementById("UI");
	if (Model.Descr.Type == pfcCreate("pfcModelType").MDL_PART || Model.GetParam("SE_MATERIAL_1").Value.StringValue == "HUA17334" ) {
		var modelHoles = Model.ListFeaturesByType (false, pfcCreate ("pfcFeatureType").FEATTYPE_HOLE);
		Model.GetParam("SE_MATERIAL_1").Value.StringValue;
		var textNode = document.createTextNode(modelHoles.Count + " holes found in the model to process.");
		UI.appendChild(textNode);
		UI.appendChild(document.createElement("br"));

		var holeInfos = document.createElement("div");
		// holesInfos.className = "cat";
		
		var holeInfo1 = document.createElement("span");
		holeInfo1.appendChild(document.createTextNode("Hole ID"));
		holeInfos.appendChild(holeInfo1);
		
		var holeInfo2 = document.createElement("span");
		holeInfo2.appendChild(document.createTextNode("State"));
		holeInfos.appendChild(holeInfo2);
		
		var holeInfo3 = document.createElement("span");
		holeInfo3.appendChild(document.createTextNode("Value"));
		holeInfos.appendChild(holeInfo3);

		var holeInfo4 = document.createElement("span");
		holeInfo4.appendChild(document.createTextNode("Highlight"));
		holeInfos.appendChild(holeInfo4);
		
		UI.appendChild(holeInfos);
			
		for (var i = 0; i < modelHoles.Count; i++) {
				var classType = ( i % 2  > 0 ? "odd" : "even");
				var holeInfos = document.createElement("div");
				holeInfos.className = "holeInfo "+ classType;
				var holeInfo1 = document.createElement("span");
				holeInfo1.className = "id";
				holeInfo1.id = modelHoles.Item(i).Id;

				var holeInfo2 = document.createElement("span");
				holeInfo2.className = "creation";

				var holeInfo3 = document.createElement("span");
				var holeInfo3Input = document.createElement("input");
				holeInfo3Input.type= "text";

				var holeInfo4 = document.createElement("span");
				var holeInfo4input = document.createElement("input");
				holeInfo4input.type= "button";
				holeInfo4input.value = "highlight";
				holeInfo4input.onclick = "alert(modelHoles.Item(i).Id);";

				holeInfo1.appendChild(document.createTextNode( modelHoles.Item(i).Id ));
				
			if (modelHoles.Item(i).GetParam("ECAD_HOLE_TYPE") == void null) {
				holeInfo2.appendChild(document.createTextNode( "created" ));
				holeInfo3Input.value= "NPTH";
				modelHoles.Item(i).CreateParam("ECAD_HOLE_TYPE", createParamValueFromString("NPTH"));
			} else {
				holeInfo2.appendChild(document.createTextNode( "current" ));
				holeInfo3Input.value= modelHoles.Item(i).GetParam("ECAD_HOLE_TYPE").Value.StringValue;
				}
			holeInfo3.appendChild(holeInfo3Input);
			holeInfo4.appendChild(holeInfo4input);
			
			holeInfos.appendChild(holeInfo1);
			holeInfos.appendChild(holeInfo2);
			holeInfos.appendChild(holeInfo3);
			holeInfos.appendChild(holeInfo4);
			UI.appendChild(holeInfos);
		}
		oSession.GetModelWindow (Model).Repaint(); 
	} else {
		var textNode = document.createTextNode("Seriously?");
		UI.appendChild(textNode);
	}
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

	
