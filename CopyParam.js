    //Provisional on=auto off=forced off
    var Provisional = "off"
    
    //By
    var ByParams = new Array("SE_DESIGNER","SE_APPROVER");
    var ByValue = new Array(DesignerName,AppoverName);
    
    //Rev
    var RevParams = new Array("SE_REVISION");
    var RevModelValue = new Array(DrawingRev);
    var RevDwgValue = new Array(DrawingRev);
    
    //list of Copy Params SETUP
    var ReqParams = new Array("SE_DESCRIPTION_ENGLISH"   ,"SE_DESCRIPTION_LOCAL"     ,"SE_PROJECT"   ,"PDM_SERVER"   ,"SE_NOTE"  );
    var ReqValue =  new Array(""                         ,""                         ,ProjectName    ,"none"         ,""         );
    
    //Initialization Commands
    var mGlob = pfcCreate("MpfcCOMGlobal");
    var oSession = mGlob.GetProESession();
    //This gets the drawing
    var Asm = oSession.CurrentModel;
    //This gets the drawing
    var Dwg = oSession.CurrentModel;
    //This variable hold list of drawing models
    var DwgModels;
    //Note users may try to use this application outside
    //of it's context; you need to check the input
    var Output = "<ul>";
    var outputTable = "";
    
    //Material
    var Material;
    //Models
    var part;
    var assembly;
        
    function CopyParam()
    {
        //top.window.external.ptc("dd");
        
        if (Asm.Descr.Type == pfcCreate("pfcModelType").MDL_ASSEMBLY) {
            //SE_PART_NUMBER Feat Parameter
            var FeatComponants = oSession.GetModel(Asm.InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY).ListFeaturesByType(false, pfcCreate("pfcFeatureType").FEATTYPE_COMPONENT);
            for (var i=0;i<FeatComponants.Count;i++)
            {
                if (FeatComponants.Item(i).GetParam("SE_PART_NUMBER") == void null) {
                    FeatComponants.Item(i).CreateParam("SE_PART_NUMBER", createParamValueFromString(FeatComponants.Item(i).ModelDescr.InstanceName));
                    Output += "<li>" + "SE_PART_NUMBER" + " : " + FeatComponants.Item(i).GetParam("SE_PART_NUMBER").Value.StringValue + " (create in assembly)</li>";
                }
            }
            
            //Final result
            if (Output == "<ul>") {
                Output += "All done, Nothing to do !</ul>";
            } else {
                //document.getElementById("application").style.display = "none";
                //document.getElementById("warning").style.display = "none";
                Output += "Done !</ul>";
            }
    
            //Refresh UI
            document.getElementById("UI").innerHTML += Output;
            
            //Regen Assembly
            oSession.RunMacro("~ Command `ProCmdRegenAuto`");
            
        } else if (Dwg.Descr.Type == pfcCreate("pfcModelType").MDL_DRAWING) {
        //Create list of all models on drawing
            DwgModels = Dwg.ListModels();
            //Build set of links to click
    
            if (DwgModels.Count > 0) {
                
                //Define part and assembly
                                            assembly = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY);
                if (assembly == void null)  assembly = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY);
                
                                            part = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART);
                if (part == void null)      part = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART);
                if (part == void null)      part = oSession.GetModel(assembly.InstanceName,pfcCreate("pfcModelType").MDL_PART);

                //Table Parameters Header
                outputTable += "<table>";
                outputTable += "<thead><tr><th scope='col'>Parameters</th><th scope='col'>values</th></tr></thead>";
                outputTable += "<tbody>";
    
                //By
                for (var i=0; i<ByParams.length; i++)
                {
                    if (Dwg.GetParam(ByParams[i]).Value.StringValue != ByValue[i]) {
                        Dwg.GetParam(ByParams[i]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(ByValue[i]);
                        Output += "<li>" + ByParams[i] + " : " + ByValue[i] + "</li>";
                    }
                    WarningMsg(ByParams[i]);
                    outputTable += "<tr class='" + "odd" +"'><td>" + ByParams[i] + "</td><td><input type='text' onChange='UpdateValue(this.name,this.value)' name ='" + ByParams[i] + "' value='" + Dwg.GetParam(ByParams[i]).Value.StringValue + "' ></input></td></tr>";
                }
    
                //Date
                var DateParams = new Array("SE_DATE");
                if (Dwg.GetParam(DateParams[0]).Value.StringValue == false ) {
                    var Today = new Date();
                    var Today = (Today.getDate() < 10 ? '0' : '') +  Today.getDate() + "/" + (Today.getMonth() + 1 < 10 ? '0' : '') + (Today.getMonth() + 1) + "/" + Today.getFullYear();
                    Dwg.GetParam(DateParams[0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(Today);
                    Output += "<li>" + DateParams[0] + " : " + Today + "</li>";
                }
                outputTable += "<tr class='" + "odd" +"'><td>" + DateParams[0] + "</td><td><input type='text' onChange='UpdateValue(this.name,this.value)' name ='" + DateParams[0] + "' value='" + Dwg.GetParam(DateParams[0]).Value.StringValue + "' ></input></td></tr>";
                WarningMsg(DateParams[0]);
                
                //Drawing Rev
                if (Dwg.GetParam(RevParams[0]).Value.StringValue == false ) {
                    if (Dwg.GetParam(RevParams[0]).Value.StringValue == "") {
                        Dwg.GetParam(RevParams[0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(RevDwgValue[0]);
                        Output += "<li>" + RevParams[0] + " : " + RevDwgValue[0] + " (drw rev with default value)</li>";
                    }
                    // WarningMsg(DateParams[0]);
                }
                outputTable += "<tr class='" + "odd" +"'><td>" + RevParams[0] + "</td><td><input type='text' name ='" + RevParams[0] + "' value='" + Dwg.GetParam(RevParams[0]).Value.StringValue + "' ></input></td></tr>";

                //Models Rev
                if (part != void null) {
                    if (part.GetParam(RevParams[0]).Value.StringValue == false ) {
                        if (part.GetParam(RevParams[0]).Value.StringValue == "") {
                            part.GetParam(RevParams[0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(RevModelValue[0]);
                            Output += "<li>" + RevParams[0] + " : " + RevModelValue[0] + " (prt rev with default value)</li>";
                        }
                        // WarningMsg(DateParams[0]);
                    }
                    outputTable += "<tr class='" + "odd" +"'><td>" + RevParams[0] + "</td><td><input type='text' name ='" + RevParams[0] + "' value='" + part.GetParam(RevParams[0]).Value.StringValue + "' ></input></td></tr>";
                }
                if (assembly != void null) {
                    if (assembly.GetParam(RevParams[0]).Value.StringValue == false ) {
                        if (assembly.GetParam(RevParams[0]).Value.StringValue == "") {
                            assembly.GetParam(RevParams[0]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(RevModelValue[0]);
                            Output += "<li>" + RevParams[0] + " : " + RevModelValue[0] + " (asm rev with default value)</li>";
                        }
                        // WarningMsg(DateParams[0]);
                    }
                    outputTable += "<tr class='" + "odd" +"'><td>" + RevParams[0] + "</td><td><input type='text' name ='" + RevParams[0] + "' value='" + assembly.GetParam(RevParams[0]).Value.StringValue + "' ></input></td></tr>";
                }
    
                //Synchronize all ReqParams Value
                writeValues(0);
                
                //Relalations
                if (part != void null)      PrtRelation(part);
                if (assembly != void null)  AsmRelation(assembly);
                
                //Regen material relations
                if (part != void null)      part.RegeneratePostRegenerationRelations();
                if (assembly != void null)  assembly.RegeneratePostRegenerationRelations();
                
                //Material
                if (DwgModels.Item(0).Descr.Type == pfcCreate("pfcModelType").MDL_PART && DwgModels.Item(0).GetParam("SE_MATERIAL_1").Value.StringValue.length > 0) {
                    //SE_Material_1 in PRT
                    Material = DwgModels.Item(0).GetParam("SE_MATERIAL_1").Value.StringValue;
    
                } else if (DwgModels.Item(0).Descr.Type == pfcCreate("pfcModelType").MDL_ASSEMBLY && oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam("SE_MATERIAL_1").Value.StringValue.length > 0) {
                    //SE_Material_1 in ASM
                    Material = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam("SE_MATERIAL_1").Value.StringValue;
    
                    //ReWrite Table Function
                    ReWriteTable();
                    
                } else {
                    warning.innerHTML += "Material in PRT not found<br />";
                    document.getElementById("fieldset1").style.border = "2px solid red";
                    document.getElementById("warning").style.color = "red";
                }
                outputTable += "<tr class='" + "odd" +"'><td>" + "SE_MATERIAL_1" + "</td><td><input type='text' onChange='UpdateValue(this.name,this.value)' name ='" + "SE_MATERIAL_1" + "' value='" + Material + "' ></input></td></tr>";
                WarningMsg("SE_MATERIAL_1");

                //Tolerance
                var TolParams = new Array("SE_TOLERANCING","SE_GENERAL_TOL");
                if (Material.length > 0) {
                    var Name = DwgModels.Item(0).InstanceName;
                    // ARTICLE TYPE 880 || COPPER ALU || COPPER CU 250 || COPPER CU 240
                    if (Name.indexOf("880-") > -1 || Material.indexOf("1CPR011961") > -1 || Material.indexOf("HUA10488") > -1 || Material.indexOf("1CPR013231") > -1) {
                        var TolValue = new Array("0S-GTS-BUSBARS","");
                        writeDwgValue(0,TolParams[0],TolValue[0]);
                        writeDwgValue(0,TolParams[1],TolValue[1]);
    
                        //Parent PRT Session ID
                        var SID=oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART).RelationId
                        var CopperNoteParams = new Array("NOTES:" ,"- THICKNESS: &SMT_THICKNESS:"+SID+"[.1]" ,"- BEND RADIUS: &SMT_THICKNESS:"+SID+"[.1]" ,"  FULL ROUND EDGE BUSBAR" ,"  UNLESS SPECIFIED OTHERWISE");
                        
                        //Create Tickness Note
                        //createSurfNote(CopperNoteParams);
                        
                    // ARTICLE TYPE 870 || GALVA || EZ || ALU || STAINLESS
                    } else if (Name.indexOf("870-") > -1 || Material.indexOf("HUA10316") > -1 || Material.indexOf("1STE010897") > -1 || Material.indexOf("HUA12808") > -1 || Material.indexOf("1STE009949") > -1) {
                        var TolValue = new Array("0S-GTS-SHEETM","");
                        writeDwgValue(0,TolParams[0],TolValue[0]);
                        writeDwgValue(0,TolParams[1],TolValue[1]);
                    // ARTICLE TYPE 876 || PC || PVC || PMMA || MAKROLON
                    } else if (Name.indexOf("876-") > -1 || Material.indexOf("1TPM009040") > -1 || Material.indexOf("HUA35575") > -1 || Material.indexOf("1TPM006917") > -1 || Material.indexOf("HUA19947") > -1) {
                        var TolValue = new Array("0S-GTS-PASTSHEET","");
                        writeDwgValue(0,TolParams[0],TolValue[0]);
                        writeDwgValue(0,TolParams[1],TolValue[1]);
                    }
                }
                WarningMsg(TolParams[0]);
                outputTable += "<tr class='" + "odd" +"'><td>" + TolParams[0] + "</td><td><input type='text' onChange='UpdateValue(this.name,this.value)' name ='" + TolParams[0] + "' value='" + Dwg.GetParam(TolParams[0]).Value.StringValue + "' ></input></td></tr>";
                outputTable += "<tr class='" + "odd" +"'><td>" + TolParams[1] + "</td><td><input type='text' onChange='UpdateValue(this.name,this.value)' name ='" + TolParams[1] + "' value='" + Dwg.GetParam(TolParams[1]).Value.StringValue + "' ></input></td></tr>";
                
                //SE_PART_NUMBER Feat Parameter
                if (DwgModels.Item(0).Descr.Type == pfcCreate("pfcModelType").MDL_ASSEMBLY) {
                    var FeatComponants = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY).ListFeaturesByType(false, pfcCreate("pfcFeatureType").FEATTYPE_COMPONENT);
                    for (var i=0;i<FeatComponants.Count;i++)
                    {
                        if (FeatComponants.Item(i).GetParam("SE_PART_NUMBER") == void null) {
                            FeatComponants.Item(i).CreateParam("SE_PART_NUMBER", createParamValueFromString(FeatComponants.Item(i).ModelDescr.InstanceName));
                            Output += "<li>" + "SE_PART_NUMBER" + " : " + FeatComponants.Item(i).GetParam("SE_PART_NUMBER").Value.StringValue + " (create in assembly)</li>";
                        }
                        //WarningMsg(DateParams[0]);
                    }
                }
    
            } else {
                document.getElementById("warning").innerHTML +=  DwgModels.Count + " models found in the drawing. Not compliant with PIM rules<br />";
                document.getElementById("fieldset1").style.border = "2px solid red";
            }
    
            //Final result
            if (Output == "<ul>") {
                Output += "All done, Nothing to do !</ul>";
            } else {
                //document.getElementById("application").style.display = "none";
                //document.getElementById("warning").style.display = "none";
                Output += "Done !</ul>";
            }
    
            //Refresh UI
            document.getElementById("UI").innerHTML += Output;
            
            outputTable += "</tbody></table>";
            document.getElementById("ManualUpdateParameters").innerHTML = "<fieldset><legend>Manual Configuration - " + Dwg.InstanceName  + " : "  + oSession.GetModel(part.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam("SE_DESCRIPTION_ENGLISH").Value.StringValue + "</legend>" + outputTable + "</fieldset>";

            //Refresh Dawing Sheets,Tables,Draft and repaint
            oSession.RunMacro("~ Command `ProCmdDwgUpdateSheets`");
            //oSession.RunMacro("~ Command `ProCmdDwgRegenDraft`");
            oSession.RunMacro("~ Command `ProCmdDwgTblRegUpd`");
            //oSession.RunMacro("~ Command `ProCmdViewRepaint`");
            oSession.CurrentWindow.Repaint();
            
            //Regen Models
            //oSession.RunMacro("~ Command `ProCmdDwgRegenModel`");
            //oSession.RunMacro("#AUTOMATIC");
        } else {
        //Inform user of problem
            warning.innerHTML="YOU CAN ONLY RUN THIS APPLICATION ON A DRAWING";
            document.getElementById("fieldset1").style.border = "2px solid red";
            document.getElementById("warning").style.color = "red";
            //document.getElementById("warning").style.border = "solid red 1px";
        }
    }
    
    //Program Functions
    //-----------------
    function writeValues(i) {
        for (k=0;k<ReqParams.length;k++)
        {
            if (Dwg.InstanceName.indexOf(DwgModels.Item(0).InstanceName) > -1) {
                if (DwgModels.Item(0).Descr.Type == pfcCreate("pfcModelType").MDL_PART) {
                    // DRW PARAM SYNC FROM MODEL
                    var PrtModel = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART)

                    //Create Parameter
                    if (PrtModel.GetParam(ReqParams[k]) == void null)               { PrtModel.CreateParam(ReqParams[k], createParamValueFromString("")); }
                    if (Dwg.GetParam(ReqParams[k]) == void null)                    { Dwg.CreateParam(ReqParams[k], createParamValueFromString("")); }

                    //Set Default Value
                    if (PrtModel.GetParam(ReqParams[k]).Value.StringValue.length == 0   && ReqValue[k].length > 0)  { PrtModel.GetParam(ReqParams[k]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(ReqValue[k]);   Output += "<li>" + ReqParams[k] + " : " + ReqValue[k] + " (prt sync with default value)</li>"; }
                    if (Dwg.GetParam(ReqParams[k]).Value.StringValue.length == 0        && ReqValue[k].length > 0)  { Dwg.GetParam(ReqParams[k]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(ReqValue[k]);        Output += "<li>" + ReqParams[k] + " : " + ReqValue[k] + " (drw sync with default value)</li>"; }

                    //DRW PARAM SYNC FROM PRT
                    if (Dwg.GetParam(ReqParams[k]).Value.StringValue != DwgModels.Item(i).GetParam(ReqParams[k]).Value.StringValue) {
                        Dwg.GetParam(ReqParams[k]).Value = DwgModels.Item(i).GetParam(ReqParams[k]).Value;
                        Output += "<li>" + ReqParams[k] + " : " + Dwg.GetParam(ReqParams[k]).Value.StringValue + " (drw sync from submodel)</li>";
                    }
                } else if (DwgModels.Item(0).Descr.Type == pfcCreate("pfcModelType").MDL_ASSEMBLY) {
                    // DRW PARAM SYNC FROM MODEL AND SUBMODEL
                    var SubPrtModel = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART)
                    var SubAsmModel = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY)

                    if (SubPrtModel.InstanceName == SubAsmModel.InstanceName) {
                        //Create Parameter
                        if (SubPrtModel.GetParam(ReqParams[k]) == void null)    { SubPrtModel.CreateParam(ReqParams[k], createParamValueFromString("")); }
                        if (SubAsmModel.GetParam(ReqParams[k]) == void null)    { SubAsmModel.CreateParam(ReqParams[k], createParamValueFromString("")); }
                        if (Dwg.GetParam(ReqParams[k]) == void null)            { Dwg.CreateParam(ReqParams[k], createParamValueFromString("")); }

                        //Set Default Value
                        if (SubPrtModel.GetParam(ReqParams[k]).Value.StringValue.length == 0    && ReqValue[k].length > 0)  { SubPrtModel.GetParam(ReqParams[k]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(ReqValue[k]);    Output += "<li>" + ReqParams[k] + " : " + ReqValue[k] + " (prt sync with default value)</li>"; }
                        if (SubAsmModel.GetParam(ReqParams[k]).Value.StringValue.length == 0    && ReqValue[k].length > 0)  { SubAsmModel.GetParam(ReqParams[k]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(ReqValue[k]);    Output += "<li>" + ReqParams[k] + " : " + ReqValue[k] + " (asm sync with default value)</li>"; }
                        if (Dwg.GetParam(ReqParams[k]).Value.StringValue.length == 0            && ReqValue[k].length > 0)  { Dwg.GetParam(ReqParams[k]).Value = pfcCreate("MpfcModelItem").CreateStringParamValue(ReqValue[k]);            Output += "<li>" + ReqParams[k] + " : " + ReqValue[k] + " (drw sync with default value)</li>"; }

                        //DRW PARAM SYNC FROM PRT
                        if (Dwg.GetParam(ReqParams[k]).Value.StringValue != SubPrtModel.GetParam(ReqParams[k]).Value.StringValue) {
                            Dwg.GetParam(ReqParams[k]).Value = SubPrtModel.GetParam(ReqParams[k]).Value;
                            Output += "<li>" + ReqParams[k] + " : " + Dwg.GetParam(ReqParams[k]).Value.StringValue + " (drw sync from prt)</li>";
                        }
                        //ASM PARAM SYNC FROM PRT
                        if (SubAsmModel.GetParam(ReqParams[k]).Value.StringValue != SubPrtModel.GetParam(ReqParams[k]).Value.StringValue) {
                            SubAsmModel.GetParam(ReqParams[k]).Value = SubPrtModel.GetParam(ReqParams[k]).Value;
                            Output += "<li>" + ReqParams[k] + " : " + Dwg.GetParam(ReqParams[k]).Value.StringValue + " (asm sync from prt)</li>";
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
            WarningMsg(ReqParams[k]);
            outputTable += "<tr class='" + "odd" +"'><td>" + ReqParams[k] + "</td><td><input type='text' onChange='UpdateValue(this.name,this.value)' name ='" + ReqParams[k] + "' value='" + Dwg.GetParam(ReqParams[k]).Value.StringValue + "' ></input></td></tr>";
        }
        //document.getElementById("application").style.display = "none";
        //document.getElementById("warning").style.display = "none";
        //document.getElementById("UI").innerHTML = "Done !<br />" + DwgModels.Item(i).InstanceName + " : Parameters values copied";
        //oSession.RunMacro("~ Command `ProCmdDwgTblRegUpd`");
    }

    function writeDwgValue(i,j,k) {
        if (Dwg.GetParam(j).Value.StringValue == false && Dwg.GetParam(j).Value.StringValue != k) {
            //write defaut parm to dwg
            Dwg.GetParam(j).Value =                      pfcCreate("MpfcModelItem").CreateStringParamValue(k);
            Output += "<li>" + j + " : " + k + "</li>";
        }
    }

    function writeDefaultValue(i,j,k) {
        if (Dwg.InstanceName == DwgModels.Item(i).InstanceName) {
                if (DwgModels.Item(0).Descr.Type == pfcCreate("pfcModelType").MDL_PART) {
                    if (Dwg.GetParam(j).Value.StringValue.length == 0 && (DwgModels.Item(i).GetParam(j).Value.StringValue != k || Dwg.GetParam(j).Value.StringValue != k)) {
                        //write defaut parm to model
                        DwgModels.Item(i).GetParam(j).Value =       pfcCreate("MpfcModelItem").CreateStringParamValue(k);
                        //write defaut parm to dwg
                        Dwg.GetParam(j).Value =                     pfcCreate("MpfcModelItem").CreateStringParamValue(k);
                        Output += "<li>" + j + " : " + k + " (sync from default value)</li>";
                    }
                } else if (DwgModels.Item(0).Descr.Type == pfcCreate("pfcModelType").MDL_ASSEMBLY) {
                    var SubPrtModel = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART)
                    var SubAsmModel = oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY)

                    if (SubPrtModel.InstanceName == SubAsmModel.InstanceName) {
                        if (Dwg.GetParam(j).Value.StringValue.length == 0 && (SubAsmModel.GetParam(j).Value.StringValue != k || SubPrtModel.GetParam(j).Value.StringValue != k || Dwg.GetParam(j).Value.StringValue != k)) {
                            //write defaut parm to asm
                            SubAsmModel.GetParam(j).Value =         pfcCreate("MpfcModelItem").CreateStringParamValue(k);
                            //write defaut parm to prt
                            SubPrtModel.GetParam(j).Value =         pfcCreate("MpfcModelItem").CreateStringParamValue(k);
                            //write defaut parm to dwg
                            Dwg.GetParam(j).Value =                 pfcCreate("MpfcModelItem").CreateStringParamValue(k);
                            Output += "<li>" + j + " : " + k + " (sync from default value)</li>";
                        }
                    }
                }
        }
    }
    
    function ReWriteTable() {
        //Parent PRT Session ID
        var SID=oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART).RelationId;
        
        var tables = Dwg.ListTables();
        for (var j=0; j<tables.Count; j++)
        {
            var table = tables.Item(j);
        
            //Build a matrix containing the values for the table
            var nTableRows = table.GetRowCount();
            var nTableCols = table.GetColumnCount();
            //alert("Table index: " + j + " Row:" + nTableRows + " Col:" + nTableCols);
        
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
                            // alert(Out + " " + nTableRows + " " + nTableCols);
                            var Released = true;
                                //var ReleaseParam Dwg.GetParam("PTC_WM_LIFECYCLE_STATE");
                                if (Dwg.GetParam("PTC_WM_LIFECYCLE_STATE") == void null) { Released = false; }      else { if (Dwg.GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue != "Released") Released = false; }
                            if (part != void null) {
                                if (part.GetParam("PTC_WM_LIFECYCLE_STATE") == void null) { Released = false; }     else { if (part.GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue != "Released") Released = false; }
                            }
                            if (assembly != void null) {
                                if (assembly.GetParam("PTC_WM_LIFECYCLE_STATE") == void null) { Released = false; } else { if (assembly.GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue != "Released") Released = false; }
                            }
    
                            //alert(Released);

                            if (Provisional != "on") {
                                ModifyCellText(table, Cell, "", 6.0, "filled"); Output += "<li>Drawing State" + " : " + "Provisional" + "</li>";
                            }
                            else
                            if (Out != "Provisional" && !Released)
                            {
                                ModifyCellText(table, Cell, "Provisional", 6.0, "filled"); Output += "<li>Drawing State" + " : " + "Provisional" + "</li>";
                            }
                            //alert(oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue);
                            //alert(DwgModels.Item(0).GetParam("PTC_WM_LIFECYCLE_STATE").Value.StringValue);
                        }
                        //alert("Table index: " + j + " Row:" + k+"/"+nTableRows+ " Col:" + l+"/"+nTableCols+ " Val:" + Out + " Val2:" + Val + " count:" + Val.Count); // + " note:" + Note);
                    }
                    catch(er)
                    {
                        // Analyse only on CellNote
                        if (table.GetCellNote(Cell) != null) {
                            if (table.GetCellNote(Cell).GetInstructions(false).TextLines.Count) {
                                if (MaterialTable == table) {
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
                //alert("/"+CellNoteTextLineText+"/"+ValXL+"/");
                //if (CellNoteTextLineText === ValXL) return true
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
            var Found_se_material_1=1;
            var Found_se_mass=0;
            var Found_se_volume=0;
            var Found_se_surface=0;

            for (var j=0;j<assembly.PostRegenerationRelations.Count; j++)
            {
                if (assembly.PostRegenerationRelations.item(j).indexOf("se_material_1") > -1)   Found_se_material_1=0;
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
                //alert(Found_se_mass+" "+Found_se_volume+" "+Found_se_surface);
            }
        }
    }

    function PrtRelation(part) {
        //var part = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART);
        if (part.RelationId > -1) {
            var Found_se_material_1=0;
            var Found_se_mass=0;
            var Found_se_volume=0;
            var Found_se_surface=0;

            for (var j=0;j<part.PostRegenerationRelations.Count; j++)
            {
                if (part.PostRegenerationRelations.item(j).indexOf("se_material_1") > -1)   Found_se_material_1=1;
                if (part.PostRegenerationRelations.item(j).indexOf("se_mass") > -1)         Found_se_mass=1;
                if (part.PostRegenerationRelations.item(j).indexOf("se_volume") > -1)       Found_se_volume=1;
                if (part.PostRegenerationRelations.item(j).indexOf("se_surface") > -1)  Found_se_surface=1;
            }

            if (Found_se_material_1==0 || Found_se_mass==0 || Found_se_volume==0 || Found_se_surface==0) {
                var relations = pfcCreate("stringseq");
                relations.Append("/* Value \"condition\" in file material");
                relations.Append("se_material_1=ptc_material_name+\"-\"+material_param(\"condition\")");
                relations.Append("/* Value SE Mass parameter with Creo Parameter");
                relations.Append("se_mass=pro_mp_mass*1e6");
                relations.Append("/* Value SE Volume parameter with Creo Parameter");
                relations.Append("se_volume=pro_mp_volume/1000");
                relations.Append("/* Value SE Surface parameter with Creo Parameter");
                relations.Append("se_surface=pro_mp_area/100");

                part.PostRegenerationRelations = relations;
                part.RegeneratePostRegenerationRelations();
                Output += "<li>Relation in prt : Updated</li>";
                //alert(Found_se_material_1+" "+Found_se_mass+" "+Found_se_volume+" "+Found_se_surface);
            }
        }
    }

    function WarningMsg(DwgParam) {
        if (DwgParam.indexOf("SE_DESCRIPTION_LOCAL") > -1 || DwgParam.indexOf("SE_NOTE") > -1) return true

        if (DwgParam.indexOf("SE_MATERIAL") > -1 || DwgParam.indexOf("SE_TREATMENT") > -1) {
            //alert(oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam(DwgParam).Value.StringValue+" "+oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam(DwgParam).Value.StringValue.indexOf("undef"));
            if (oSession.GetModel(part.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam(DwgParam).Value.StringValue.length == 0 || oSession.GetModel(part.InstanceName,pfcCreate("pfcModelType").MDL_PART).GetParam(DwgParam).Value.StringValue.indexOf("undef") > -1) {
                /*if (warning.innerHTML.indexOf("Copying model parameters") > -1) {
                    warning.innerHTML="";
                }*/
                warning.innerHTML += DwgParam + " not specified<br />";
                document.getElementById("fieldset1").style.border = "2px solid red";
                document.getElementById("warning").style.color = "red";
            }
        } else if (Dwg.GetParam(DwgParam).Value.StringValue.length == 0) {
            /*if (warning.innerHTML.indexOf("Copying model parameters") > -1) {
                warning.innerHTML="";
            }*/
            warning.innerHTML += DwgParam + " not specified<br />";
            document.getElementById("fieldset1").style.border = "2px solid red";
            document.getElementById("warning").style.color = "red";
        }
    }
    
    function UpdateValue(name,value) {
        //alert(name+" "+value);
        var SubPrtModel = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_PART);
        var SubAsmModel = oSession.GetModel(Dwg.InstanceName,pfcCreate("pfcModelType").MDL_ASSEMBLY);

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
        if (Dwg != void null) {
            //Create DRW Param if not found
            if (Dwg.GetParam(name) == void null)            { Dwg.CreateParam(name, createParamValueFromString("")); }
            //Set new value
            Dwg.GetParam(name).Value =          pfcCreate("MpfcModelItem").CreateStringParamValue(value);
        }
        //alert("Param Updated:"+name+" Value:"+value);
        
        oSession.RunMacro("~ Command `ProCmdDwgUpdateSheets`");
        oSession.RunMacro("~ Command `ProCmdDwgTblRegUpd`");
        oSession.CurrentWindow.Repaint();
    }
    
    function AssignMaterial() {
        server= oSession.GetActiveServer();
        
        var pfccheckoutoptions = pfcCreate ("pfcCheckoutOptions"); 
        var checkoutoptions =  pfccheckoutoptions.Create(); 
        checkoutoptions.Download =true;
        
        try {
            var partfilename = "hua10316.mtl";
            var partname = "hua10316";
            var wslink = server.CheckoutObjects(null,partfilename,false,checkoutoptions);
        }catch(e) {
            alert("Unable to find " + partfilename + " in MDM");
        }  	
        
        oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART).RetrieveMaterial(partname);
        
        //alert(oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART).ListMaterials().Item(0).Name);
        //alert(oSession.GetModel(DwgModels.Item(0).InstanceName,pfcCreate("pfcModelType").MDL_PART).CurrentMaterial.GetPropertyValue(Name));
    }