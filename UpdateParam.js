    //list of Update Params SETUP
    var updatedParameters = new Array ("SE_DESIGNER","SE_APPROVER", "SE_DESCRIPTION_LOCAL","SE_REVISION",   "SE_PROJECT",   "SE_NOTE")
    var updatedParametersValues = new Array ("AMEG JCA","BMA","--" , "00","GALAXY VL", "--" )

    //Initialization Commands
    var mGlob = pfcCreate("MpfcCOMGlobal");
    var oSession = mGlob.GetProESession();
    var currentModel = oSession.CurrentModel;
    var currentModelParametersList = new Array ();
    
    var Today = new Date();
    var seDateFound = false;
    var seParameterFound = false;
    
    var Today = new Date();
    Today = (Today.getDate() < 10 ? '0' : '') +  Today.getDate() + "/" + (Today.getMonth() + 1 < 10 ? '0' : '') + (Today.getMonth() + 1) + "/" + Today.getFullYear();

    function UpdateParam() {
        var oddEven = "odd";
        getParameterList();

        for ( var listedParameter = 0; listedParameter < currentModelParametersList.length; listedParameter++) {
            if (currentModelParametersList[listedParameter] == "SE_DATE") {
                seDateFound = true;
            }
        }
        if (!seDateFound) {
            currentModel.CreateParam("SE_DATE", CreateParamValue("STRING", Today));
        } else {
            if (currentModel.GetParam("SE_DATE").Value.StringValue == "*") {
                alert(Today);
                currentModel.GetParam("SE_DATE").Value.StringValue = Today;
            }
        }
        for ( var k = 0 ;  k < updatedParameters.length; k++) {
            for ( var listedParameter = 0; listedParameter < currentModelParametersList.length; listedParameter++) {
                if (updatedParameters[k] == currentModelParametersList[listedParameter]) {
                    seParameterFound = true;
                }
            }
            if ( seParameterFound == false) {
                currentModel.CreateParam(updatedParameters[k], CreateParamValue("STRING", updatedParametersValues[k]));
            }
            seParameterFound = false;
        }
        var parameterStringValue = currentModel.GetParam("SE_DESCRIPTION_ENGLISH").Value.StringValue;
        var descriptionEnglish = (parameterStringValue == "" ? "No description set !" : parameterStringValue );
        var outputTable = "";
        outputTable += "<table><caption>" + currentModel.InstanceName  + " : "  + descriptionEnglish + "</caption>";
        outputTable += "<thead><tr><th scope='col'>Parameters</th><th scope='col'>3D model values</th><th>Copy </th></th><th scope='col'>2D model values</th></tr></thead>";
        outputTable += "<tbody>";
        outputTable += "<tr class='even'><td>SE_DATE</td>";
        outputTable += "<td>"+  currentModel.GetParam("SE_DATE").Value.StringValue + "</td>";
        outputTable += "<td><input type='checkbox' /></td>";
        outputTable += "<td><input type='text' name ='SE_DATE' value='" + Today + "' ></input></td></tr>";
        outputTable += "</tr>";

        for (k = 0; k < updatedParameters.length; k++) {
            parameterStringValue = currentModel.GetParam(updatedParameters[k]).Value.StringValue;
            outputTable += "<tr class='" + oddEven +"'><td>" + updatedParameters[k] + "</td>";
            outputTable += "<td>" + (parameterStringValue == "" ? "Not set !" : parameterStringValue) + "</td>";
            outputTable += "<td><input type='checkbox' /></td>";
            outputTable += "<td><input type='text' name ='" + updatedParameters[k] + "' value='" + updatedParametersValues[k] + "' ></input></td></tr>";
            oddEven = (k % 2 ? "odd":"even");

        }
        outputTable += "</tbody></table>";
        
        document.getElementById("updatedParameters").innerHTML += "<fieldset><legend>Configuration</legend>" + outputTable + "</fieldset>";
    }

    //Program Functions
    function getParameterList() {
        for ( var k = 0; k < currentModel.ListParams().Count; k++) {
            currentModelParametersList.push(currentModel.ListParams().item(k).Name);
        }
    }
    function CreateParamValue(typeAsString,val) {
        switch (typeAsString) {
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
    function pfcIsWindows () {
        return (navigator.appName.indexOf ("Microsoft") != -1 ? true : false)
    }
    function pfcCreate (className)  {
        if (!pfcIsWindows()) {
            netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
        }

        return (pfcIsWindows() ? new ActiveXObject ("pfc."+className) : Components.classes ["@ptc.com/pfc/" + className + ";1"].createInstance());
    }