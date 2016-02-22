    //Provisional on=auto off=forced off
    var Provisional = "off"

    var Today = new Date();
    var Today = (Today.getDate() < 10 ? '0' : '') +  Today.getDate() + "/" + (Today.getMonth() + 1 < 10 ? '0' : '') + (Today.getMonth() + 1) + "/" + Today.getFullYear();

    //Params manager
    var Params = new Array();
    //                     SE_PARAM                  SYNC        !Value      Value      Def Drw Value   Def Model Value
    Params[0] = new Array("SE_DESIGNER",            "nosync",   "default",  "default",  DesignerName,   "null");
    Params[1] = new Array("SE_APPROVER",            "nosync",   "default",  "manual",   AppoverName,    "null");
    Params[2] = new Array("SE_DATE",                "nosync",   "default",  "default",  Today,          "null");
    Params[3] = new Array("SE_REVISION",            "nosync",   "default",  "manual",   DrawingRev,     ModelsRev);
    Params[4] = new Array("SE_DESCRIPTION_ENGLISH", "sync",     "manual",   "manual",   "",             "");
    Params[5] = new Array("SE_DESCRIPTION_LOCAL",   "sync",     "manual",   "manual",   "",             "");
    Params[6] = new Array("SE_PROJECT",             "sync",     "default",  "manual",   ProjectName,    ProjectName);
    Params[7] = new Array("PDM_SERVER",             "sync",     "default",  "default",  "none",         "none");
    Params[8] = new Array("SE_NOTE",                "sync",     "manual",   "manual",   "",             "");

    //Initialization Commands
    var mGlob = pfcCreate("MpfcCOMGlobal");
    var oSession = mGlob.GetProESession();
    var CurrentModel = oSession.CurrentModel;

    //This set all current models type
    var Asm = CurrentModel; var Prt = CurrentModel; var Dwg = CurrentModel;

    //This variable hold list of drawing models
    var DwgModels;
    var modelsName;
    //Note users may try to use this application outside
    //of it's context; you need to check the input
    var Output = "<ul>";
    var outputTable = "";

    //Material
    var Material;

    //Models
    var part;
    var assembly;

    //materialList for sub function
    var matList
    if (materialList != void null) matList = materialList;

    function CopyParam()
    {
        if (CurrentModel == void null) {
            //Inform user of problem
            warning.innerHTML="YOU CAN ONLY RUN THIS APPLICATION ON MODEL";
            document.getElementById("fieldset1").style.border = "2px solid red";
            document.getElementById("warning").style.color = "red";
            
        } else if (Prt.Descr.Type == pfcCreate("pfcModelType").MDL_PART) {

            modelsName = Prt.InstanceName;

            part = oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART);

            AssignMaterial();
        /*
            //Holes highlight
            oSession.CurrentWindow.Repaint();
            var FeatComponants = oSession.GetModel(Prt.InstanceName,pfcCreate("pfcModelType").MDL_PART).ListFeaturesByType(false, pfcCreate("pfcFeatureType").FEATTYPE_HOLE);
            for (var i=0;i<FeatComponants.Count;i++)
            {
                pfcCreate("MpfcSelect").CreateModelItemSelection(CurrentModel.GetItemById(pfcCreate("pfcModelItemType").ITEM_FEATURE, FeatComponants.Item(i).Id), null).Display();
            }

            //Final result
            if (Output == "<ul>") {
                Output += "Hole features highlighted</ul>";
            } else {
                Output += "</ul>";
            }

            //Refresh UI
            document.getElementById("UI").innerHTML += Output;

            //Regen Assembly
            oSession.RunMacro("~ Command `ProCmdRegenAuto`");
        */
        } else if (Asm.Descr.Type == pfcCreate("pfcModelType").MDL_ASSEMBLY) {
            //SE_PART_NUMBER Feat Parameter
            var FeatComponants = oSession.GetModel(Asm.InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY).ListFeaturesByType(false, pfcCreate("pfcFeatureType").FEATTYPE_COMPONENT);
            for (var i=0;i<FeatComponants.Count;i++)
            {
                if (FeatComponants.Item(i).GetParam("SE_PART_NUMBER") == void null) {
                    FeatComponants.Item(i).CreateParam("SE_PART_NUMBER", CreateParamValue("STRING",FeatComponants.Item(i).ModelDescr.InstanceName));
                    Output += "<li>" + "SE_PART_NUMBER" + " : " + FeatComponants.Item(i).GetParam("SE_PART_NUMBER").Value.StringValue + " (create in assembly)</li>";
                }
            }

            //Final result
            if (Output == "<ul>") {
                Output += "All done, Nothing to do !</ul>";
            } else {
                Output += "Done !</ul>";
            }

            //Refresh UI
            document.getElementById("UI").innerHTML += Output;

            //Regen Assembly
            oSession.RunMacro("~ Command `ProCmdRegenAuto`");

        } else if (Dwg.Descr.Type == pfcCreate("pfcModelType").MDL_DRAWING) {
            try {
                // Browser window default size
                var browserSize = oSession.CurrentWindow.GetBrowserSize();
                if (browserSize > 35.0) {
                    browserSize = 35.0;
                    oSession.CurrentWindow.SetBrowserSize(0.0);
                    oSession.CurrentWindow.SetBrowserSize(browserSize);
                }
            } catch(e) {}
            
            //Create list of all models on drawing
            DwgModels = Dwg.ListModels();
            //Build set of links to click
            if (DwgModels.Count > 0) {
                //Defind models name
                modelsName = DwgModels.Item(0).InstanceName;

                //Define part and assembly
                                            assembly = oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_ASSEMBLY);
                if (assembly == void null)  assembly = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY);

                                            part = oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART);
                if (part == void null)      part = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART);
                if (part == void null)      part = oSession.GetModel(assembly.InstanceName,pfcCreate("pfcModelType").MDL_PART);

                if (! matList.length) {
                    // 3PP ETO default List
                    matList = [
                        new Array("HUA10488","COPPER BAR ALU"),
                        new Array("1CPR011961","COPPER BAR CU 250"),
                        new Array("HUA10316","GALVA"),
                        new Array("1STE010897","EZ"),
                        new Array("HUA12808","ALU"),
                        new Array("1STE009949","STAINLESS"),
                        new Array("1TPM009040","PC"),
                        new Array("HUA35575","PVC"),
                        new Array("1TPM006917","PMMA"),
                        new Array("HUA19947","MAKROLON")
                    ];
                }

                //Table Parameters Header
                outputTable += "<table>";
                outputTable += "<thead><tr><th scope='col'>Parameters</th><th scope='col'>default</th><th scope='col' colspan='2' rowspan='1'>values</th></tr></thead>";
                outputTable += "<tbody>";

                //Synchronize all Params Value
                for (i=0;i<Params.length && syncValues(i);i++) { }

                //Relalations
                if (part != void null)      PrtRelation(part);
                if (assembly != void null)  AsmRelation(assembly);

                //Regen material relations
                try {
                    if (part != void null)      part.RegeneratePostRegenerationRelations();
                    if (assembly != void null)  assembly.RegeneratePostRegenerationRelations();
                } catch(er) {
                    warning.innerHTML += "Relation error, can be regenerated<br />";
                    document.getElementById("fieldset1").style.border = "2px solid red";
                    document.getElementById("warning").style.color = "red";
                }

                //Material
                if (assembly == void null && part != void null && DwgModels.Item(0).GetParam("SE_MATERIAL_1").Value.StringValue.length > 0) {
                    //SE_Material_1 in PRT
                    Material = DwgModels.Item(0).GetParam("SE_MATERIAL_1").Value.StringValue;
                    
                    //ReWrite Table Function
                    ReWriteTable();

                } else if (assembly != void null && part != void null && oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART).GetParam("SE_MATERIAL_1").Value.StringValue.length > 0) {
                    //SE_Material_1 in ASM
                    Material = oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART).GetParam("SE_MATERIAL_1").Value.StringValue;

                    //ReWrite Table Function
                    ReWriteTable();

                } else {
                    warning.innerHTML += "Material in PRT not found<br />";
                    document.getElementById("fieldset1").style.border = "2px solid red";
                    document.getElementById("warning").style.color = "red";
                }
                outputTable += "<tr class='" + "odd" +"'><td class='odd1'>" + "SE_MATERIAL_1" + "</td><td class='odd2'></td><td colspan='2' rowspan='1'><input type='text' id='Input_NoSync' name ='" + "SE_MATERIAL_1" + "' value='" + Material + "' ></input></td></tr>";
                WarningMsg("SE_MATERIAL_1");

                //Tolerance
                var TolParams = new Array("SE_TOLERANCING","SE_GENERAL_TOL");
                if (Material != void null) {
                    // ARTICLE TYPE 880 || COPPER ALU || COPPER CU 250 || COPPER CU 240
                    if (modelsName.indexOf("880-") > -1 || Material.indexOf("HUA10488") > -1 || Material.indexOf("1CPR011961") > -1 || Material.indexOf("1CPR013231") > -1) {
                        var TolValue = new Array("0S-GTS-BUSBARS","");
                        writeDwgValue(TolParams[0],TolValue[0]);
                        writeDwgValue(TolParams[1],TolValue[1]);

                        //Parent PRT Session ID
                        var SID=oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART).RelationId
                        var CopperNoteParams = new Array("NOTES:" ,"- THICKNESS: &SMT_THICKNESS:"+SID+"[.1]" ,"- BEND RADIUS: &SMT_THICKNESS:"+SID+"[.1]" ,"  FULL ROUND EDGE BUSBAR" ,"  UNLESS SPECIFIED OTHERWISE");

                        //Create Tickness Note
                        //createSurfNote(CopperNoteParams);

                    // ARTICLE TYPE 870 || GALVA || EZ || ALU || STAINLESS
                    } else if (modelsName.indexOf("870-") > -1 || Material.indexOf("HUA10316") > -1 || Material.indexOf("1STE010897") > -1 || Material.indexOf("HUA12808") > -1 || Material.indexOf("1STE009949") > -1) {
                        var TolValue = new Array("0S-GTS-SHEETM","");
                        writeDwgValue(TolParams[0],TolValue[0]);
                        writeDwgValue(TolParams[1],TolValue[1]);
                    // ARTICLE TYPE 876 || PC || PVC || PMMA || MAKROLON
                    } else if (modelsName.indexOf("876-") > -1 || Material.indexOf("1TPM009040") > -1 || Material.indexOf("HUA35575") > -1 || Material.indexOf("1TPM006917") > -1 || Material.indexOf("HUA19947") > -1) {
                        var TolValue = new Array("0S-GTS-PASTSHEET","");
                        writeDwgValue(TolParams[0],TolValue[0]);
                        writeDwgValue(TolParams[1],TolValue[1]);
                    }
                }

                WarningMsg(TolParams[0]);
                outputTable += "<tr class='" + "odd" +"'><td class='odd1'>" + TolParams[0] + "</td><td class='odd2'></td><td colspan='2' rowspan='1'><input type='text' id='Input_NoSync' onChange='UpdateValue(this.name,this.value,this.id)' name ='" + TolParams[0] + "' value='" + Dwg.GetParam(TolParams[0]).Value.StringValue + "' ></input></td></tr>";
                outputTable += "<tr class='" + "odd" +"'><td class='odd1'>" + TolParams[1] + "</td><td class='odd2'></td><td colspan='2' rowspan='1'><input type='text' id='Input_NoSync' onChange='UpdateValue(this.name,this.value,this.id)' name ='" + TolParams[1] + "' value='" + Dwg.GetParam(TolParams[1]).Value.StringValue + "' ></input></td></tr>";

                //SE_PART_NUMBER Feat Parameter
                if (assembly != void null) {
                    var FeatComponants = assembly.ListFeaturesByType(false, pfcCreate("pfcFeatureType").FEATTYPE_COMPONENT);
                    for (var i=0;i<FeatComponants.Count;i++)
                    {
                        var Featname = FeatComponants.Item(i).ModelDescr.GetFullName();
                        if (FeatComponants.Item(i).GetParam("SE_PART_NUMBER") == void null) {
                            FeatComponants.Item(i).CreateParam("SE_PART_NUMBER", CreateParamValue("STRING",String(Featname)));
                            Output += "<li>" + "SE_PART_NUMBER" + " : " + String(FeatComponants.Item(i).GetParam("SE_PART_NUMBER").Value.StringValue) + " (create in assembly)</li>";
                        }
                    }
                }

                //Create material Select List
                outputTable += "<tr class='" + "odd" +"'><td class='odd1'>" + "MATERIAL" + "</td><td class='odd2'></td><td colspan='2' rowspan='1'>";
                outputTable += "<select name='select' id='Select_Def' onchange='AssignMaterial(this)'>";
                var curMat
                if (part.CurrentMaterial) curMat = part.CurrentMaterial.Name;
                if (curMat == void null || curMat.toUpperCase() == "UNDEF") {
                    outputTable += "<option selected></option>"
                } else {
                    outputTable += "<option></option>"
                }

                for (var i=0;i<matList.length;i++) {
                    var space = "";
                    if ( matList[i][0].length < 10 ) { for (j=0;(10 - j - matList[i][0].length)>0;j++) { var space = space + "&nbsp" } }

                    if (curMat != void null && curMat.toUpperCase() == matList[i][0].toUpperCase()) {
                        outputTable += "<option value='"+ matList[i][0] +"' selected>" + matList[i][0] + space + " - "+ matList[i][1] + "</option>"
                    } else {
                        outputTable += "<option value='"+ matList[i][0] +"'>" + matList[i][0] + space + " - "+ matList[i][1] + "</option>"
                    }
                }
                outputTable += "</select>";
                outputTable += "</td></tr>";

            } else {
                document.getElementById("warning").innerHTML +=  DwgModels.Count + " models found in the drawing. Not compliant with PIM rules<br />";
                document.getElementById("fieldset1").style.border = "2px solid red";
            }

            //Final result
            if (Output == "<ul>") {
                Output += "All done, Nothing to do !</ul>";
            } else {
                Output += "Done !</ul>";
            }

            //Refresh UI
            document.getElementById("UI").innerHTML += Output;

            outputTable += "</tbody></table>";

            document.getElementById("ManualUpdateParameters").innerHTML = "<fieldset id='fieldset'><legend>Manual Configuration - " + Dwg.InstanceName  + " : "  + oSession.GetModel(part.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam("SE_DESCRIPTION_ENGLISH").Value.StringValue + "</legend>" + outputTable + "</fieldset>";

            //Refresh Dawing Sheets,Tables,Draft and repaint
            oSession.RunMacro("~ Command `ProCmdDwgUpdateSheets`");
            //oSession.RunMacro("~ Command `ProCmdDwgRegenDraft`");
            oSession.RunMacro("~ Command `ProCmdDwgTblRegUpd`");
            //oSession.RunMacro("~ Command `ProCmdViewRepaint`");
            oSession.CurrentWindow.Repaint();

            //pfcRegenInstructions.RefreshModelTree

            //Regen Models
            //oSession.RunMacro("~ Command `ProCmdDwgRegenModel`");
            //oSession.RunMacro("#AUTOMATIC");
        } else {
        //Inform user of problem
            warning.innerHTML="YOU CAN ONLY RUN THIS APPLICATION ON A DRAWING";
            document.getElementById("fieldset1").style.border = "2px solid red";
            document.getElementById("warning").style.color = "red";
        }
    }

    //Program Functions
    //-----------------
    function syncValues(i) {
        if (Dwg.InstanceName.indexOf(modelsName) > -1) {
            if (assembly == void null && part != void null) {
                // DRW PARAM SYNC FROM MODEL
                var PrtModel = oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART);
                //Create Parameter
                if (PrtModel.GetParam(Params[i][0]) == void null    && Params[i][5] != "null")  { PrtModel.CreateParam(Params[i][0], createParamValueFromString("")); }
                if (Dwg.GetParam(Params[i][0]) == void null         && Params[i][4] != "null")  { Dwg.CreateParam(Params[i][0], createParamValueFromString("")); }
                //Set Default Value
                if (Params[i][5] != "null" && PrtModel.GetParam(Params[i][0]).Value.StringValue != Params[i][5]) { if ((PrtModel.GetParam(Params[i][0]).Value.StringValue.length == 0   && Params[i][2] == "default") || (PrtModel.GetParam(Params[i][0]).Value.StringValue.length > 0   && Params[i][3] == "default")) { PrtModel.GetParam(Params[i][0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(Params[i][5]);   Output += "<li>" + Params[i][0] + " : " + Params[i][5] + " (prt sync with default value)</li>"; } }
                if (Params[i][4] != "null" && Dwg.GetParam(Params[i][0]).Value.StringValue != Params[i][4])      { if ((Dwg.GetParam(Params[i][0]).Value.StringValue.length == 0        && Params[i][2] == "default") || (Dwg.GetParam(Params[i][0]).Value.StringValue.length > 0        && Params[i][3] == "default")) { Dwg.GetParam(Params[i][0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(Params[i][4]);        Output += "<li>" + Params[i][0] + " : " + Params[i][4] + " (drw sync with default value)</li>"; } }
                //DRW PARAM SYNC FROM PRT
                if (Params[i][1] == "sync") {
                    if (Dwg.GetParam(Params[i][0]).Value.StringValue != DwgModels.Item(0).GetParam(Params[i][0]).Value.StringValue) {
                        Dwg.GetParam(Params[i][0]).Value = DwgModels.Item(0).GetParam(Params[i][0]).Value;
                        Output += "<li>" + Params[i][0] + " : " + Dwg.GetParam(Params[i][0]).Value.StringValue + " (drw sync from submodel)</li>";
                    }
                }
            } else if (assembly != void null && part != void null) {
                // DRW PARAM SYNC FROM MODEL AND SUBMODEL
                var SubPrtModel = oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART);
                var SubAsmModel = oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_ASSEMBLY);
                if (SubPrtModel.InstanceName == SubAsmModel.InstanceName) {
                    //Create Parameter
                    if (SubPrtModel.GetParam(Params[i][0]) == void null &&  Params[i][5] != "null")     { SubPrtModel.CreateParam(Params[i][0], createParamValueFromString("")); }
                    if (SubAsmModel.GetParam(Params[i][0]) == void null &&  Params[i][5] != "null")     { SubAsmModel.CreateParam(Params[i][0], createParamValueFromString("")); }
                    if (Dwg.GetParam(Params[i][0]) == void null         &&  Params[i][4] != "null")     { Dwg.CreateParam(Params[i][0], createParamValueFromString("")); }
                    //Set Default Value
                    if (Params[i][5] != "null" && SubPrtModel.GetParam(Params[i][0]).Value.StringValue != Params[i][5])  { if ((SubPrtModel.GetParam(Params[i][0]).Value.StringValue.length == 0    && Params[i][2] == "default") || (SubPrtModel.GetParam(Params[i][0]).Value.StringValue.length > 0    && Params[i][3] == "default"))  { SubPrtModel.GetParam(Params[i][0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(Params[i][5]);    Output += "<li>" + Params[i][0] + " : " + Params[i][5] + " (prt sync with default value)</li>"; } }
                    if (Params[i][5] != "null" && SubAsmModel.GetParam(Params[i][0]).Value.StringValue != Params[i][5])  { if ((SubAsmModel.GetParam(Params[i][0]).Value.StringValue.length == 0    && Params[i][2] == "default") || (SubAsmModel.GetParam(Params[i][0]).Value.StringValue.length > 0    && Params[i][3] == "default"))  { SubAsmModel.GetParam(Params[i][0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(Params[i][5]);    Output += "<li>" + Params[i][0] + " : " + Params[i][5] + " (asm sync with default value)</li>"; } }
                    if (Params[i][4] != "null" && Dwg.GetParam(Params[i][0]).Value.StringValue != Params[i][4])          { if ((Dwg.GetParam(Params[i][0]).Value.StringValue.length == 0            && Params[i][2] == "default") || (Dwg.GetParam(Params[i][0]).Value.StringValue.length > 0            && Params[i][3] == "default"))  { Dwg.GetParam(Params[i][0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(Params[i][4]);            Output += "<li>" + Params[i][0] + " : " + Params[i][4] + " (drw sync with default value)</li>"; } }
                    //DRW PARAM SYNC FROM PRT
                    if (Params[i][1] == "sync") {
                        if (Dwg.GetParam(Params[i][0]).Value.StringValue != SubPrtModel.GetParam(Params[i][0]).Value.StringValue) {
                            Dwg.GetParam(Params[i][0]).Value = SubPrtModel.GetParam(Params[i][0]).Value;
                            Output += "<li>" + Params[i][0] + " : " + Dwg.GetParam(Params[i][0]).Value.StringValue + " (drw sync from prt)</li>";
                        }
                    }
                    //ASM PARAM SYNC FROM PRT
                    if (Params[i][5] != "null") {
                        if (SubAsmModel.GetParam(Params[i][0]).Value.StringValue != SubPrtModel.GetParam(Params[i][0]).Value.StringValue) {
                            SubAsmModel.GetParam(Params[i][0]).Value = SubPrtModel.GetParam(Params[i][0]).Value;
                            Output += "<li>" + Params[i][0] + " : " + Dwg.GetParam(Params[i][0]).Value.StringValue + " (asm sync from prt)</li>";
                        }
                    }
                } else {
                    warning.innerHTML += "Model and sub Model do not have the same name. Parameter need to be set manually<br />";
                    document.getElementById("fieldset1").style.border = "2px solid red";
                    return false;
                }
            }
        } else {
            warning.innerHTML += "Drawing and Model do not have the same name. Parameter need to be set manually<br />";
            document.getElementById("fieldset1").style.border = "2px solid red";
            return false;
        }

        WarningMsg(Params[i][0]);
        if (Params[i][1] != "sync" && Params[i][5] != "null") {
            outputTable += "<tr class='" + "odd" +"'><td class='odd1'>" + Params[i][0] + "</td><td>";
            if (Params[i][4].length) { outputTable += "<button type='button' class='button' id='Input_Drw' onclick='UpdateValue(this.name,this.value,this.id,this.type)' name ='" + Params[i][0] + "' value='" + Params[i][4] + "'>"+ Params[i][4] +"</button><br>"; }
            if (Params[i][5].length) { outputTable += "<button type='button' class='button' id='Input_Models' onclick='UpdateValue(this.name,this.value,this.id,this.type)' name ='" + Params[i][0] + "' value='" + Params[i][5] + "'>"+ Params[i][5] +"</button>"; }
            outputTable += "</td><td colspan='2' rowspan='1'><input type='text' id='Input_Drw' onkeypress='handleKeyPress(event,this.name,this.value,this.id)' onChange='UpdateValue(this.name,this.value,this.id)' name ='" + Params[i][0] + "' value='" + Dwg.GetParam(Params[i][0]).Value.StringValue + "' ></input>";
            outputTable += "<input type='text' id='Input_Models' onChange='UpdateValue(this.name,this.value,this.id)' name ='" + Params[i][0] + "' value='" + oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART).GetParam(Params[i][0]).Value.StringValue + "' ></input></td></tr>";
        } else {
            if (Params[i][4].length) { outputTable += "<tr class='" + "odd" +"'><td class='odd1'>" + Params[i][0] + "</td>"; }
            else { outputTable += "<tr class='" + "odd" +"'><td class='odd1' colspan='2' rowspan='1'>" + Params[i][0] + "</td>"; }
            if (Params[i][4].length) { outputTable += "<td><button type='button' id='Input_Drw' onclick='UpdateValue(this.name,this.value,this.id,this.type)' name ='" + Params[i][0] + "' value='" + Params[i][4] + "'>"+ Params[i][4] +"</button></td>"; }
            outputTable += "<td colspan='2' rowspan='1'><input type='text' id='Input_"+Params[i][1]+"' onkeypress='handleKeyPress(event,this.name,this.value,this.id)' onChange='UpdateValue(this.name,this.value,this.id)' name ='" + Params[i][0] + "' value='" + Dwg.GetParam(Params[i][0]).Value.StringValue + "' ></input></td></tr>";
        }
        
        return true;
    }

    function writeDwgValue(j,k) {
        if (Dwg.GetParam(j) == void null)    { Dwg.CreateParam(j, createParamValueFromString("")); }
        if (Dwg.GetParam(j).Value.StringValue != k) {
            //write defaut parm to dwg
            Dwg.GetParam(j).Value =                      pfcCreate("MpfcModelItem").CreateStringParamValue(k);
            if (! k.length) k = "clear parameter";
            Output += "<li>" + j + " : " + k + "</li>";
        }
    }

    function ReWriteTable() {
        //Parent PRT Session ID
        var SID=oSession.GetModel(modelsName,pfcCreate("pfcModelType").MDL_PART).RelationId;

        var tables = Dwg.ListTables();
        for (var j=0; j<tables.Count; j++)
        {
            var table = tables.Item(j);

            //Build a matrix containing the values for the table
            var nTableRows = table.GetRowCount();
            var nTableCols = table.GetColumnCount();

            //Loop around the table and dump information to excel
            for (k=0;k<nTableRows;k++)
            {
                for (l=0;l<nTableCols;l++)
                {
                    var Cell = pfcCreate("pfcTableCell").Create(k+1, l+1);
                    var Mode = pfcCreate("pfcParamMode").DWGTABLE_NORMAL; // DWGTABLE_FULL {0:VALUE}
                    var Out="";

                    if (MaterialTable != table) var MaterialTable = null;

                    try
                    {
                        var Val = table.GetText(Cell,Mode);

                        if (Val.Count>1) { MultipleLinesInCells = true; }

                        for (m=0;m<Val.Count;m++)
                        {
                            if (m>0) { if (m<Val.Count) { Out = Out + " "; } }
                            Out = Out + Val.Item(k);
                        }

                        //Find Material Table
                        if (Out.indexOf("Part range or family") > -1) {
                            var MaterialTable = table;
                        }

                        //Provisional Analysis
                        if (nTableRows==1 && nTableCols==1) {
                            var Released = true;
                                if (Dwg.GetParam("PTC_WM_LIFECYCLE_STATE") == void null) { Released = false; }      else { if (Dwg.GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue != "Released") Released = false; }
                            if (part != void null) {
                                if (part.GetParam("PTC_WM_LIFECYCLE_STATE") == void null) { Released = false; }     else { if (part.GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue != "Released") Released = false; }
                            }
                            if (assembly != void null) {
                                if (assembly.GetParam("PTC_WM_LIFECYCLE_STATE") == void null) { Released = false; } else { if (assembly.GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue != "Released") Released = false; }
                            }

                            if (Provisional != "on") {
                                if (Out.indexOf("Provisional") > -1) {
                                    ModifyCellText(table, Cell, "", 6.0, "filled"); Output += "<li>Drawing State" + " : " + "Provisional" + "(Removed)</li>";
                                }
                            }
                            else
                            if (Out != "Provisional" && !Released)
                            {
                                ModifyCellText(table, Cell, "Provisional", 6.0, "filled"); Output += "<li>Drawing State" + " : " + "Provisional" + "</li>";
                            }
                        }
                    }
                    catch(er)
                    {
                        // Analyse only on CellNote
                        if (table.GetCellNote(Cell) != null) {
                            if (table.GetCellNote(Cell).GetInstructions(false).TextLines.Count) {
                                if (MaterialTable == table && assembly != void null) {
                                    //Set right col
                                    var CellPramMat = pfcCreate("pfcTableCell").Create(k+1, l+2);
                                    if (FindCellText(table, Cell, "Material 1")) { if(!FindCellText(table, CellPramMat, "&SE_MATERIAL_1:"+SID)) { ModifyCellText(table, CellPramMat, " &SE_MATERIAL_1:"+SID, 2.1, "font"); Output += "<li>SE_MATERIAL_1" + " : " + "SE_MATERIAL_1:"+SID + "</li>"; } }
                                    if (FindCellText(table, Cell, "Material 2")) { if(!FindCellText(table, CellPramMat, "&SE_MATERIAL_2:"+SID)) { ModifyCellText(table, CellPramMat, " &SE_MATERIAL_2:"+SID, 2.1, "font"); Output += "<li>SE_MATERIAL_2" + " : " + "SE_MATERIAL_2:"+SID + "</li>"; } }
                                    if (FindCellText(table, Cell, "Material 3")) { if(!FindCellText(table, CellPramMat, "&SE_MATERIAL_3:"+SID)) { ModifyCellText(table, CellPramMat, " &SE_MATERIAL_3:"+SID, 2.1, "font"); Output += "<li>SE_MATERIAL_3" + " : " + "SE_MATERIAL_3:"+SID + "</li>"; } }
                                    if (FindCellText(table, Cell, "Treatment 1")) { if(!FindCellText(table, CellPramMat, "&SE_TREATMENT_1:"+SID)) { ModifyCellText(table, CellPramMat, " &SE_TREATMENT_1:"+SID, 2.1, "font"); Output += "<li>SE_TREATMENT_1" + " : " + "SE_TREATMENT_1:"+SID + "</li>"; } }
                                    if (FindCellText(table, Cell, "Treatment 2")) { if(!FindCellText(table, CellPramMat, "&SE_TREATMENT_2:"+SID)) { ModifyCellText(table, CellPramMat, " &SE_TREATMENT_2:"+SID, 2.1, "font"); Output += "<li>SE_TREATMENT_2" + " : " + "SE_TREATMENT_2:"+SID + "</li>"; } }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    //Main function
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
        }
        else {
            obj = Components.classes ["@ptc.com/pfc/" + className + ";1"].createInstance();
        }

        // Return created object
        //
        if (className.match(/^M?pfc/)) {
            global_class_cache[className] = obj;
        }

        return obj;
    }
    //Utility Functions
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

    /*====================================================================*\
    FUNCTION: createParamValueFromString
    PURPOSE:  Parses a string into a pfcParamValue object, checking for most
            restrictive possible type to use.
    \*====================================================================*/
    function createParamValueFromString (s /* string */)
    {
        if (s.indexOf (".") == -1)
        {
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

    /* FUNCTION: FindCellText */
    function FindCellText(Table, cell, ValXL)
    {
        //var lines = pfcCreate("stringseq");
        //lines.Append (ValXL);
        //Table.SetText(cell, lines);

        //var drawing = oSession.CurrentModel;
        //var matrix = drawing.GetSheetTransform(drawing.CurrentSheetNumber).Matrix;

        var CellNote = Table.GetCellNote(cell);
        var CellNoteInsts = CellNote.GetInstructions(true);
        var CellNoteTextLines=CellNoteInsts.TextLines;
        for(var i=0;i<CellNoteTextLines.Count;i++)
        {
            var CellNoteTextLine=CellNoteTextLines.Item(i);
            var CellNoteTextLineTexts=CellNoteTextLine.Texts;

            for(var j=0;j<CellNoteTextLineTexts.Count;j++)
            {
                var CellNoteTextLineText=CellNoteTextLineTexts.Item(j).Text;
                if (CellNoteTextLineText.indexOf(ValXL) > -1) return true
                //CellNoteTextLineText.TextHeight=2.1/(matrix.Item(0, 0));//<-need to transform the coordinates
            }
        }

        //CellNote.Modify(CellNoteInsts);
        return false
    }
    /* FUNCTION: ModifyCellText */
    function ModifyCellText(Table, cell, ValXL, Size, Font)
    {
        //First get the current font from the cell
        var CellNote = Table.GetCellNote(cell);
        var CellNoteInsts = CellNote.GetInstructions(true);
        var CellNoteTextLines = CellNoteInsts.TextLines;
        var CellNoteTextLine1 = CellNoteInsts.TextLines.Item(0);
        var CellNoteTextLine1Texts = CellNoteTextLine1.Texts;
        var CellNoteTextLine1Text1 = CellNoteTextLine1Texts.Item(0);
        var FontName = CellNoteTextLine1Text1.FontName;

        var lines = pfcCreate("stringseq");
        lines.Append (ValXL);
        Table.SetText(cell, lines);

        var drawing = oSession.CurrentModel;
        var matrix = drawing.GetSheetTransform(drawing.CurrentSheetNumber).Matrix;

        var CellNote = Table.GetCellNote(cell);
        var CellNoteInsts = CellNote.GetInstructions(true);
        var CellNoteTextLines = CellNoteInsts.TextLines;

        for(var i=0;i<CellNoteTextLines.Count;i++)
        {
            var CellNoteTextLine = CellNoteTextLines.Item(i);
            var CellNoteTextLineTexts = CellNoteTextLine.Texts;

            for(var j=0;j<CellNoteTextLineTexts.Count;j++)
            {
                var CellNoteTextLineText = CellNoteTextLineTexts.Item(j);
                CellNoteTextLineText.TextHeight = Size/(matrix.Item(0, 0));//<-need to transform the coordinates
                CellNoteTextLineText.FontName = FontName;
                // CellNoteTextLineText.FontName = Font;
                // CellNoteTextLineText.WidthFactor = 1.0; // 0.6 / 1
            }
        }

        CellNote.Modify(CellNoteInsts);
    }
    /* FUNCTION: writeTextInCell */
    function writeTextInCell(table /* pfcTable */, row /* integer */,col /* integer */, text /* string */)
     {
        var cell = pfcCreate ("pfcTableCell").Create (row, col);
        var lines = pfcCreate ("stringseq");
        lines.Append (text);

        try{
        table.SetText (cell, lines);
        }
        catch(er)
        {
            alert (row+"  "+col);
        }
     }
    /*====================================================================*\
    FUNCTION : createSurfNote()
    PURPOSE  : Utility to create a note that documents the surface name or id.
    The note text will be placed at the upper right corner of the selected view.
    \*====================================================================*/
    function createSurfNote(NoteArray)
    {

    /*--------------------------------------------------------------------*\
      Get the current drawing & its background view
    \*--------------------------------------------------------------------*/
      var session = oSession;
      var drawing = session.CurrentModel;

    //  if (drawing.Type != pfcCreate("pfcModelType").MDL_DRAWING)
    //    throw new Error(0, "Current model is not a drawing");

    /*--------------------------------------------------------------------*\
      Interactively select a surface in a drawing view
    \*--------------------------------------------------------------------*/
    /*
    alert("dd");
      var browserSize = session.CurrentWindow.GetBrowserSize();
      session.CurrentWindow.SetBrowserSize(0.0);
      alert("dd");
      var options = pfcCreate ("pfcSelectionOptions").Create ("surface");
      options.MaxNumSels = 1;
      alert("dd");
      var sels = session.Select (options, void null);
      var selSurf = sels.Item (0);
      var item = selSurf.SelItem;
      alert("dd");
      var name = item.GetName();
      if (name == void null)
        name = new String ("Surface ID "+item.Id);
      alert("dd");
      session.CurrentWindow.SetBrowserSize(browserSize);
        alert("dd");
       */
        //var name = "";
        //name = new String ("NOTES:");
        //name += new String ("- THICKNESS: &SMT_THICKNESS:0");

        //var NoteParams = new Array("NOTES:" ,"- THICKNESS: &SMT_THICKNESS:"+sid ,"- BEND RADIUS: &SMT_THICKNESS:"+sid ,"  FULL ROUND EDGE BUSBAR" ,"  UNLESS SPECIFIED OTHERWISE");
        //var i=0;
        //alert(NoteParams.length);
        for (i=0;i<NoteArray.length;i++)
        {
            if (i==0) {
                var text = pfcCreate("pfcDetailText").Create(NoteArray[i]);
                var texts = pfcCreate("pfcDetailTexts");
            } else {
                text = pfcCreate("pfcDetailText").Create(NoteArray[i]);
                //Clear Texts
                texts = pfcCreate("pfcDetailTexts");
            }

            texts.Append(text);

            if (i==0) {
                var textLine = pfcCreate("pfcDetailTextLine").Create(texts);
                var textLines = pfcCreate("pfcDetailTextLines");
            } else {
                textLine = pfcCreate("pfcDetailTextLine").Create(texts);
                //textLines = pfcCreate("pfcDetailTextLines");
            }

            textLines.Append(textLine);
        }
        //- BEND RADIUS: &SMT_THICKNESS:0  FULL ROUND EDGE BUSBAR  UNLESS SPECIFIED OTHERWISE");

    /*--------------------------------------------------------------------*\
      Allocate a text item, and set its properties
    \*--------------------------------------------------------------------*/
      //var text = pfcCreate ("pfcDetailText").Create (name);

    /*--------------------------------------------------------------------*\
      Allocate a new text line, and add the text item to it
    /*--------------------------------------------------------------------*/
      //var texts = pfcCreate ("pfcDetailTexts");
      //texts.Append (text);
      /*textLines.Append ("NOTES:");
      textLines.Append ("- THICKNESS: &SMT_THICKNESS:0");
      textLines.Append ("- BEND RADIUS: &SMT_THICKNESS:0");
      textLines.Append ("  FULL ROUND EDGE BUSBAR");
      textLines.Append ("  UNLESS SPECIFIED OTHERWISE");*/

      //var textLine = pfcCreate ("pfcDetailTextLine").Create (texts);

      //var textLines = pfcCreate ("pfcDetailTextLines");
      //textLines.Append (textLine);


    /*--------------------------------------------------------------------*\
      Set the location of the note text
    \*--------------------------------------------------------------------*/

      //var dwgView = selSurf.SelView2D;
      //var outline = dwgView.Outline;
      //var textPos = outline.Item (1);

      // Force the note to be slightly beyond the view outline boundary
     //textPos.Set (0, textPos.Item (0) + 0.25 * (textPos.Item (0) -
     //                       outline.Item (0).Item(0)));
     //textPos.Set (1, textPos.Item (1) + 0.25 * (textPos.Item (1) -
     //                       outline.Item (0).Item(1)));
/*
      var dwgView = drw;
      var outline = dwgView.Outline;
      var textPos = outline.Item (1);
     textPos.Set (0, 20);
     textPos.Set (1, 20);
     var position = pfcCreate ("pfcFreeAttachment").Create (textPos);
     position.View = dwgView;
*/
     //position.View = dwgView;

    /*--------------------------------------------------------------------*\
      Set the attachment for the note leader
    \*--------------------------------------------------------------------*/
     //var leaderToSurf = pfcCreate ("pfcParametricAttachment").Create (selSurf);

    /*--------------------------------------------------------------------*\
      Set the attachment structure
    \*--------------------------------------------------------------------*/
/*
     var allAttachments = pfcCreate ("pfcDetailLeaders").Create ();
     allAttachments.ItemAttachment = position;
*/
     //allAttachments.Leaders = pfcCreate ("pfcAttachments");
     //allAttachments.Leaders.Append (leaderToSurf);

    /*--------------------------------------------------------------------*\
      Allocate a note description, and set its properties
    \*--------------------------------------------------------------------*/
     var instrs = pfcCreate ("pfcDetailNoteInstructions").Create (textLines);

//     instrs.Leader = allAttachments;

    /*--------------------------------------------------------------------*\
      Create the note
    \*--------------------------------------------------------------------*/
     var note = drawing.CreateDetailItem (instrs);

    /*--------------------------------------------------------------------*\
      Display the note
    \*--------------------------------------------------------------------*/
     note.Show ();
    }

    function AsmRelation(assembly) {
        if (assembly.RelationId > -1) {
            var Found_se_material_1=0;
            var Found_se_mass=0;
            var Found_se_volume=0;
            var Found_se_surface=0;

            for (var j=0;j<assembly.PostRegenerationRelations.Count; j++)
            {
                if (assembly.PostRegenerationRelations.item(j).indexOf("se_material_1=") > -1)  Found_se_material_1=1;
                if (assembly.PostRegenerationRelations.item(j).indexOf("se_mass") > -1)         Found_se_mass=1;
                if (assembly.PostRegenerationRelations.item(j).indexOf("se_volume") > -1)       Found_se_volume=1;
                if (assembly.PostRegenerationRelations.item(j).indexOf("se_surface") > -1)      Found_se_surface=1;
            }
            if (Found_se_material_1==1 || Found_se_mass==0 || Found_se_volume==0 || Found_se_surface==0) {
                var relations = pfcCreate("stringseq");
                relations.Append("/* Value SE Mass parameter with Creo Parameter");
                relations.Append("se_mass=pro_mp_mass*1e6");
                relations.Append("/* Value SE Volume parameter with Creo Parameter");
                relations.Append("se_volume=pro_mp_volume/1000");
                relations.Append("/* Value SE Surface parameter with Creo Parameter");
                relations.Append("se_surface=pro_mp_area/100");

                assembly.PostRegenerationRelations = relations;
                assembly.RegeneratePostRegenerationRelations();
                Output += "<li>Relation in asm : Updated</li>";
            }
        }
    }

    function PrtRelation(part) {
        if (part.RelationId > -1) {
            var Found_se_material_1=0;
            var Found_se_mass=0;
            var Found_se_volume=0;
            var Found_se_surface=0;

            for (var j=0;j<part.PostRegenerationRelations.Count; j++)
            {
                if (part.PostRegenerationRelations.item(j).indexOf("se_material_1=\"") > -1)    Found_se_material_1=1;
                if (part.PostRegenerationRelations.item(j).indexOf("se_mass") > -1)             Found_se_mass=1;
                if (part.PostRegenerationRelations.item(j).indexOf("se_volume") > -1)           Found_se_volume=1;
                if (part.PostRegenerationRelations.item(j).indexOf("se_surface") > -1)          Found_se_surface=1;
            }

            if (Found_se_material_1==0 || Found_se_mass==0 || Found_se_volume==0 || Found_se_surface==0) {
                var relations = pfcCreate("stringseq");

                // /* Value "condition" in file material
                // if (ptc_material_name == "UNASSIGNED")
                // se_material_1="UNDEF-undef"
                // else
                // se_material_1=ptc_material_name+"-"+material_param("condition")
                // endif
                // /* Value SE Mass parameter with Creo Parameter
                // se_mass=pro_mp_mass*1e6
                // /* Value SE Volume parameter with Creo Parameter
                // se_volume=pro_mp_volume/1000
                // /* Value SE Surface parameter with Creo Parameter
                // se_surface=pro_mp_area/100

                relations.Append("/* Value \"condition\" in file material");
                relations.Append("if (ptc_material_name == \"UNASSIGNED\")");
                relations.Append("se_material_1=\"UNDEF-undef\"");
                relations.Append("else");
                relations.Append("se_material_1=ptc_material_name+\"-\"+material_param(\"condition\")");
                relations.Append("endif");
                relations.Append("/* Value SE Mass parameter with Creo Parameter");
                relations.Append("se_mass=pro_mp_mass*1e6");
                relations.Append("/* Value SE Volume parameter with Creo Parameter");
                relations.Append("se_volume=pro_mp_volume/1000");
                relations.Append("/* Value SE Surface parameter with Creo Parameter");
                relations.Append("se_surface=pro_mp_area/100");

                part.PostRegenerationRelations = relations;
                part.RegeneratePostRegenerationRelations();
                Output += "<li>Relation in prt : Updated</li>";
            }
        }
    }

    function WarningMsg(DwgParam) {
        if (DwgParam.indexOf("SE_DESCRIPTION_LOCAL") > -1 || DwgParam.indexOf("SE_NOTE") > -1) return true

        if (DwgParam.indexOf("SE_MATERIAL") > -1 || DwgParam.indexOf("SE_TREATMENT") > -1) {
            if (oSession.GetModel(part.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam(DwgParam).Value.StringValue.length == 0 || oSession.GetModel(part.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam(DwgParam).Value.StringValue.indexOf("undef") > -1) {
                warning.innerHTML += DwgParam + " not specified<br />";
                document.getElementById("fieldset1").style.border = "2px solid red";
                document.getElementById("warning").style.color = "red";
            }
        } else if (Dwg.GetParam(DwgParam).Value.StringValue.length == 0) {
            warning.innerHTML += DwgParam + " not specified<br />";
            document.getElementById("fieldset1").style.border = "2px solid red";
            document.getElementById("warning").style.color = "red";
        }
    }
    
    function handleKeyPress(e,name,value,id) {
        var key=e.keyCode || e.which;
        if (key==13) {
            UpdateValue(name,value,id);
        }
    }

    function UpdateValue(name,value,id,type) {
        var SubPrtModel = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART);
        var SubAsmModel = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY);

        if (id == "Input_Sync" || id == "Input_sync" || id == "Input_Models") {
            if (SubPrtModel != void null) {
                //Create PRT Param if not found
                if (SubPrtModel.GetParam(name) == void null)    { SubPrtModel.CreateParam(name, createParamValueFromString("")); }
                //Set new value
                SubPrtModel.GetParam(name).Value =  pfcCreate("MpfcModelItem").CreateStringParamValue(value);
            }
            if (SubAsmModel != void null) {
                //Create ASM Param if not found
                if (SubAsmModel.GetParam(name) == void null)    { SubAsmModel.CreateParam(name, createParamValueFromString("")); }
                //Set new value
                SubAsmModel.GetParam(name).Value =  pfcCreate("MpfcModelItem").CreateStringParamValue(value);
            }
        }
        if (id == "Input_NoSync" || id == "Input_Sync" || id == "Input_nosync" || id == "Input_sync" || id == "Input_Drw") {
            if (Dwg != void null) {
                //Create DRW Param if not found
                if (Dwg.GetParam(name) == void null)            { Dwg.CreateParam(name, createParamValueFromString("")); }
                //Set new value
                Dwg.GetParam(name).Value =          pfcCreate("MpfcModelItem").CreateStringParamValue(value);
            }
        }

        oSession.RunMacro("~ Command `ProCmdDwgUpdateSheets`");
        oSession.RunMacro("~ Command `ProCmdDwgTblRegUpd`");
        oSession.CurrentWindow.Repaint();
        
        if (type == "button") location.reload();
    }
    
    function DefaultValue(source,target) {
        //document.getElementById(target).value = document.getElementById(source).value;
        document.getElementsByClassName(target).value = document.getElementsByClassName(source).value;
        //getElementsByTagName
    }

    function AssignMaterial(s) {
        server = oSession.GetActiveServer();

        var pfccheckoutoptions = pfcCreate ("pfcCheckoutOptions");
        var checkoutoptions =  pfccheckoutoptions.Create();
        checkoutoptions.Download =false;
        checkoutoptions.Readonly =true;

        var matFileName;
        var matName = s[s.selectedIndex].value.toLowerCase();
        if (matName == void null || matName == "" || matName.toUpperCase() == "UNDEF") matName = "undef";
        
        var matNames = pfcCreate ("stringseq");
		matNames.set (0, matName.toLowerCase() + ".mtl");
        matNames.set (1, matName.toUpperCase() + ".MTL");
        matNames.set (2, matName.toLowerCase() + ".mat");
        matNames.set (3, matName.toUpperCase() + ".MAT");

        var wsLink;
        var wsModel;
        var modelInWs;
        var materialIndex

        //Part Analysis for very fast change
        for (i=0;i<part.ListMaterials().Count;i++) {
            //find added material
            if (part.ListMaterials().Item(i).Name.toUpperCase() == matName.toUpperCase()) { materialIndex = i; }
        }

        var j = 0;
        for (i=0;i<4 && wsLink==void null && modelInWs == void null && materialIndex == void null;i++) {
            // Check if Material is on WS for quick change
            try {
                modelInWs = server.IsObjectCheckedOut(server.ActiveWorkspace, matNames.Item(i));
                if (modelInWs != void null) {
                    var matFileName = matNames.Item(i);
                } else {
                    // import Material to Session from Mdm
                    try {
                        var wsLink = server.CheckoutObjects(null,matNames.Item(i),false,checkoutoptions);
                        alert("Linked in session");
                        var matFileName = matNames.Item(i);
                    } catch(e) {
                        j++;
                        if (i >= 3) alert("Unable to find " + s[s.selectedIndex].value + " in MDM");
                    }
                }
            } catch(e) {
                // import Material to Session from Mdm
                try {
                    var wsLink = server.CheckoutObjects(null,matNames.Item(i),false,checkoutoptions);
                    alert("Linked in session");
                    var matFileName = matNames.Item(i);
                } catch(e) {
                    j++;
                    if (i >= 3) alert("Unable to find " + s[s.selectedIndex].value + " in MDM");
                }
            }
        }

        if (matFileName != void null && materialIndex == void null) {
            // Import new Matrial in part
            part.RetrieveMaterial(matName);
            //alert("added to part");
        }

        for (i=0;i<part.ListMaterials().Count;i++) {
            // Clear all part marterial
            //part.ListMaterials().Item(i).Delete();

            //find added material
            if (part.ListMaterials().Item(i).Name.toUpperCase() == matName.toUpperCase()) { materialIndex = i; }
        }

        if (materialIndex != void null) {
            // Set new Matrial in part
            part.CurrentMaterial = part.ListMaterials().Item(materialIndex);

            // alert(part.ListMaterials().Item(0).Name);
            location.reload();

            return true;
        } else {
            return false;
        }
    }

    // mdlDescr.SetPath(workdir);
    
    // Snippet code (uses functions from pfcUtils.js) :
    // Import drawing options
    /*
    function Import(){
      var session = pfcCreate("MpfcCOMGlobal").GetProESession();
      var mdl = session.CurrentModel;
      try {
     
      var dtl_inst =  pfcCreate("pfcDWGSetupImportInstructions").Create();
      mdl.Import("test.dtl", dtl_inst);
      }
      catch (err) {
        alert(err.description);
      }
    }
    */