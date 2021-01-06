using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using CivilDB = Autodesk.Civil.DatabaseServices;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using civil3dCogoPoints;

[assembly: CommandClass(typeof(CornerRecordExtract.Commands))]
[assembly: ExtensionApplication(null)]
namespace CornerRecordExtract
{
    #region Commands
    public class Commands
    {
        [CommandMethod("OCPWBR")]
        public void ListAttributes()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            var doc = Application.DocumentManager.MdiActiveDocument;

            try
            {
                var acDB = doc.Database;

                using (var trans = acDB.TransactionManager.StartTransaction())
                {
                    // Capture Layouts to be Checked
                    DBDictionary layoutPages = (DBDictionary)trans.GetObject(acDB.LayoutDictionaryId,OpenMode.ForRead);

                    // Dictionary report of corner record form
                    Dictionary<string, object> cr_dictReport = new Dictionary<string, object>();
                   
                    // Handle Corner Record meta data dictionary extracted from Properties and Content
                    Dictionary<String, object> cornerRecordForms = new Dictionary<string, object>();

                    // List of form checks
                    List<object> listOfFormChecks = new List<object>();

                    // List of Layout names 
                    List<string> layoutNamesList = new List<string>();

                    // Dictionary result of CR Form dynamic block and Layout Name.
                    Dictionary<string, object> crFormAndLayout_dictResult = new Dictionary<string, object>();

                    // Extract CogoPoints From DWG file.
                    CivilDB.CogoPointCollection cogoPointsColl = CivilDB.CogoPointCollection.GetCogoPoints(doc.Database);
                    var cogoPointCollected = CogoPointJson.geolocationCapture(cogoPointsColl);

                    foreach (DBDictionaryEntry layoutPage in layoutPages)
                    {
                        var crFormItems = layoutPage.Value.GetObject(OpenMode.ForRead) as Layout;
                        var isModelSpace = crFormItems.ModelType;

                        ObjectIdCollection textObjCollection = new ObjectIdCollection();

                        Dictionary<string, string> crAttributes = new Dictionary<string, string>();

                        // Form Check results Dictionary
                        Dictionary<string, object> formCheckResult_dict = new Dictionary<string, object>();

                        // Missing attribute information check result 
                        Dictionary<string, object> missing_attributeDataForm = new Dictionary<string, object>();

                        // Dictionary of missing "Form" error
                        Dictionary<string, object> missing_dynamicBlockForm = new Dictionary<string, object>();

                        if (isModelSpace != true)
                        {
                            BlockTableRecord blkTblRec = trans.GetObject(crFormItems.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;

                            layoutNamesList.Add(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ", ""));

                            foreach (ObjectId blkId in blkTblRec)
                            {
                                // Store object id in object in collection (TBD)
                                textObjCollection.Add(blkId);

                                var blkRef = trans.GetObject(blkId, OpenMode.ForRead) as BlockReference;

                                if (blkRef != null && blkRef.IsDynamicBlock)
                                {
                                    AttributeCollection attCol = blkRef.AttributeCollection;

                                    DynamicBlockReferencePropertyCollection testAttCol = blkRef.DynamicBlockReferencePropertyCollection;

                                    foreach (ObjectId attId in attCol)
                                    {
                                        AttributeReference attRef = (AttributeReference)trans.GetObject(attId, OpenMode.ForRead);

                                        if (attRef.Tag.ToString() == "CITY_NAME")
                                        {     
                                            if (!String.IsNullOrEmpty(attRef.TextString.ToString()))
                                            {
                                                crAttributes.Add("CRCity_c", attRef.TextString.ToString());
                                            }
                                            else
                                            {
                                                // Even though the values are missing still store the data in 
                                                // crAttributes dictionary in order to check if the 
                                                // form does exist in the current layout for further analysis.
                                                crAttributes.Add("CRCity_c", attRef.TextString.ToString());

                                                missing_attributeDataForm.Add("CRCity_c", "Could not find CRCity_c field for layout named " +
                                                    layoutPage.Key.ToString());
                                            }
                                        }

                                        if (attRef.Tag.ToString() == "LEGAL_DESCRIPTION")
                                        {
                                            if (!String.IsNullOrEmpty(attRef.TextString.ToString()))
                                            {
                                                crAttributes.Add("Corner_Type_c", "Lot");
                                                crAttributes.Add("Legal_Description_c", attRef.TextString.ToString());
                                            }
                                            else
                                            {
                                                // Even though the values are missing still store the data in 
                                                // crAttributes dictionary in order to check if the 
                                                // form does exist in the current layout for further analysis. 
                                                crAttributes.Add("Corner_Type_c", "Lot");
                                                crAttributes.Add("Legal_Description_c", attRef.TextString.ToString());

                                                missing_attributeDataForm.Add("Legal_Description_c",
                                                    "Could not find Legal_Description_c field for layout named " +
                                                    layoutPage.Key.ToString());
                                            }
                                        }
                                    }
                                }
                            }

                            // Check Layout for proper naming convention 
                            // Check if the Corner Record Form Attribute data is collected in the current layout.
                         
                            Match layoutNameMatch = Regex.Match(crFormItems.LayoutName, "^(\\s*cr\\s*\\d\\d*)$",
                                RegexOptions.IgnoreCase);
                            if (layoutNameMatch.Success)
                            {
                                if (crAttributes.ContainsKey("CRCity_c") && crAttributes.ContainsKey("Corner_Type_c")
                                && crAttributes.ContainsKey("Legal_Description_c"))
                                {
                                    // Check if the missing_attributeDataForm dictionary does not contain any keys.
                                    // If it does contian any keys then certien values in the field attribute are missing.
                                    if (missing_attributeDataForm.Count == 0)
                                    {
                                        formCheckResult_dict.Add("Form_Check_Status", "Pass");
                                        formCheckResult_dict.Add("Form_In_Layout_Error", "None");
                                        formCheckResult_dict.Add("Form_Att_Values_Error", "None");

                                        // Build Final JSON File with LayoutName and Attributes
                                        if (cornerRecordForms.ContainsKey(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ", "")))
                                        {
                                            Dictionary<string, object> duplicate_LayoutMessages = new Dictionary<string, object>();

                                            duplicate_LayoutMessages.Add(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ", "")
                                                + "_DUPLICATE", "DUPLICATE LAYOUT, CANNOT EXTRACT DATA");
                                            formCheckResult_dict.Add("Layout_Duplicate_Check", "Fail");
                                            formCheckResult_dict.Add("Layout_Duplicate_Error", duplicate_LayoutMessages);
                                        }
                                        else
                                        {
                                            cornerRecordForms.Add(crFormItems.LayoutName.Trim().ToString().ToLower().Replace(" ", ""), crAttributes);
                                        }
                                    }
                                    else
                                    {
                                        // Find a way to display the error message for missing attribute information.
                                        formCheckResult_dict.Add("Form_Check_Status", "Fail");
                                        formCheckResult_dict.Add("Form_In_Layout_Error", "None");
                                        formCheckResult_dict.Add("Form_Att_Values_Error", missing_attributeDataForm);
                                    }
                                }
                                else
                                {
                                    // Handle a way to display missing corner recoed form in a layout that is named CR #
                                    // Error_message: found a layout name {curent layout} with no corner record form
                                    // Solution_message: Rename layout or add a corner record form to the layout
                                    missing_dynamicBlockForm.Add("Error_message", "Found a layout name " + crFormItems.LayoutName +
                                        " with no corner record form");
                                    missing_dynamicBlockForm.Add("Solution_message", "Rename layout or add a corner record form to the layout");

                                    formCheckResult_dict.Add("Form_Check_Status", "Fail");
                                    formCheckResult_dict.Add("Form_In_Layout_Error", missing_dynamicBlockForm);
                                    formCheckResult_dict.Add("Form_Att_Values_Error", "None");
                                }
                            }
                        }
                        // Add Form Check Result Dictionary to List only if the results 
                        // exist to prevent model space from creating the dictionary
                        if (formCheckResult_dict.Count > 0)
                        {
                            listOfFormChecks.Add(formCheckResult_dict);
                        }
                    }

                    //if (layoutNamesList){return a list of duplicate layout names}
                    IEnumerable<string> duplicateLayouts = layoutNamesList.GroupBy(x => x).SelectMany(g => g.Skip(1));
  
                    // Checks to see whether the points from the cogo point collection exist within 
                    // the layout by searching for the correct collection key and layout name
                    List<string> cogoPointCollectedCheck = cogoPointCollected.Keys.ToList();
                    List<bool> boolCheckResults = new List<bool>();

                    // Extra Layout Dictionary Error Message
                    Dictionary<string, object> extraLayout_result = new Dictionary<string, object>();

                    // List of Extra Layers results
                    List<object> listOfExtraLayouts = new List<object>();

                    // Extra CogoPoint Dictionary Error Message
                    Dictionary<string, object> extraCogoPoint_result = new Dictionary<string, object>();

                    // List of Extra CogoPoints result
                    List<object> listOfExtraCogoPoints = new List<object>();

                    // Dictionary of Duplicate Layout Names error
                    Dictionary<string, object> duplicate_ItemMessages = new Dictionary<string, object>();

                    // In order to compaer layout names to field names from corner record point table.
                    // Check cogoPointsColl dictionary does not contain any duplicate error.
                    if (cogoPointCollected.ContainsKey("Fail_CogoPoints"))
                    {
                        cr_dictReport.Add("CogoPoint_check", cogoPointCollected);
                        //cr_dictReport.Add("Check_duplicate_name_in_tbl", cogoPointCollected);
                    }

                    if (duplicateLayouts.ToList().Any()) 
                    {
                        //cr_dictReport.Add("Check_duplicate_name_in_Layout", "Fail: " + String.Join(",", duplicateLayouts));
                        duplicate_ItemMessages.Add("Error_message", "Duplicate Layout named " + String.Join(",", duplicateLayouts));
                        duplicate_ItemMessages.Add("Solution_message", "Remove or Rename the layout");

                        extraLayout_result.Add("Check_duplicate_name_in_Layout", "Fail");
                        extraLayout_result.Add("Duplicate_Layout_Error", duplicate_ItemMessages);
                    }
                    else
                    {
                        //cr_dictReport.Add("Check_duplicate_name_in_Points", "Pass");
                        //cr_dictReport.Add("Check_duplicate_name_in_Layout", "Pass");
                        extraLayout_result.Add("Check_duplicate_name_in_Points", "Pass");
                        extraLayout_result.Add("Check_duplicate_name_in_Layout", "Pass");

                        //=======================================================================================================================
                        // DETERMINE EXTRA LAYOUT NAMES
                        // Here we are checking for corner records that are not found based on the layount name "CR#"
                        // This will return a layer name that may or may not be associated to a record form such as "Layout1" etc...
                        // So it can also mean that the layout name exist with a corner record form named "CR5" but the cogo point does not exist.
                        IEnumerable<string> cogoPointNameCheck = layoutNamesList.Except(cogoPointCollectedCheck);
                        List<string> cogoPointNameCheckResults = cogoPointNameCheck.ToList();
                        var layoutNameChecker = new Regex("^(\\s*cr\\s*\\d\\d*)$");

                        if (!cogoPointNameCheckResults.Where(f => layoutNameChecker.IsMatch(f)).ToList().Any())
                        {
                            extraLayout_result.Add("Layout_check_status", "Pass");
                            extraLayout_result.Add("Extra_Layouts", "None");
                            boolCheckResults.Add(true);
                        }
                        else
                        {
                            extraLayout_result.Add("Layout_check_status", "Fail");
                            foreach (string cogoPointNameResultItem in cogoPointNameCheckResults)
                            {
                                // Extra Layout Dictionary Error Message
                                Dictionary<string, object> storeExtraLayouts = new Dictionary<string, object>();

                                Match layoutNameMatch = Regex.Match(cogoPointNameResultItem, "^(\\s*cr\\s*\\d\\d*)$",
                                    RegexOptions.IgnoreCase);

                                if (layoutNameMatch.Success)
                                {
                                    string layoutNameX = layoutNameMatch.Value;

                                    storeExtraLayouts.Add("Extra_Layouts", "Layout Named " + layoutNameX +
                                        " does not have an associated cogo point. Please add a CogoPoint and/or update the Corner Record # column in the CogoPoint’s Table with " + layoutNameX + ". Else rename the layout " + layoutNameX + ".");
                                    listOfExtraLayouts.Add(storeExtraLayouts);
                                }
                                else
                                {
                                    //string layoutNameX = layoutNameMatch.Value;
                                    // Write message to tell users to check layouts not named CR
                                    //ed.WriteMessage(baseChecks["CogoLayoutNameMatches"]);
                                    //baseChecks["CogoLayoutNameMatches"] = "Pass";
                                }
                            }
                            // Dictionary migh change extraLayout_result (TBD) 
                            extraLayout_result.Add("Layout_Check", listOfExtraLayouts);
                            boolCheckResults.Add(false);
                        }
                        // Look into ways of better handleing erros
                        // =======================================================================================================================
                        // DETERMINE EXTRA COGO POINTS
                        // Here we are checking for extra cogo points by matching it to the layout names that are collected.
                        // We are comparing to the cogo point that exist and labeled CR in the field name.
                        IEnumerable<string> layoutNameCheck = cogoPointCollectedCheck.Except(layoutNamesList);
                        List<string> layoutNameCheckResults = layoutNameCheck.ToList();
                        var cogoNameChecker = new Regex("^(\\s*cr\\s*\\d\\d*)$");

                        // If the layout name has any value other than CR == PASS
                        // If CR point exists and does not match then throw an error for user to fix
                        if (!layoutNameCheckResults.Where(f => cogoNameChecker.IsMatch(f)).ToList().Any())
                        {
                            extraCogoPoint_result.Add("CogoPoint_check_status", "Pass");
                            extraCogoPoint_result.Add("Extra_CogoPoint", "None");
                            boolCheckResults.Add(true);
                        }
                        else // Found a CR point that DID NOT match a layout name 
                        {
                            extraCogoPoint_result.Add("CogoPoint_check_status", "Fail");
                            foreach (string layoutNameCheckResultItem in layoutNameCheckResults)
                            {
                                // Extra CogoPoint Dictionary Error Message
                                Dictionary<string, object> storeExtraCogoPoints = new Dictionary<string, object>();

                                Match cogoNameMatch = Regex.Match(layoutNameCheckResultItem, "^(\\s*cr\\s*\\d\\d*)$",
                                    RegexOptions.IgnoreCase);

                                if (cogoNameMatch.Success)
                                {
                                    string cogoNameX = cogoNameMatch.Value;

                                    storeExtraCogoPoints.Add("Extra_CogoPoint", "Corner Record point named " + cogoNameX + 
                                        " does not have an associated Layout");
                                    listOfExtraCogoPoints.Add(storeExtraCogoPoints);
                                }
                                else
                                {
                                    // Write message to tell users to check Points not named CR
                                    // Cogo point named correctly like "Monumnet P4"
                                    //ed.WriteMessage(responseMsg["Success"]);
                                }
                            }
                            // Dictionary migh change extraCogoPoint_result (TBD) 
                            extraCogoPoint_result.Add("CogoPoint_Check", listOfExtraCogoPoints);
                            boolCheckResults.Add(false);
                        }
                    }
                    
                    // ===========================================================================================================================
                    // Output JSON file to BIN folder
                    // IF there are two true booleans in the list then add the data to the corresponding keys (cr1 => cr1)
                    if ((boolCheckResults.Count(v => v == true)) == 2)
                    {
                        foreach (string cornerRecordFormKey in cornerRecordForms.Keys)
                        {
                            if (cogoPointCollected.ContainsKey(cornerRecordFormKey))
                            {
                                var cogoFinal = (Dictionary<string, string>)cornerRecordForms[cornerRecordFormKey];
                                var cogoFinalLong = ((Dictionary<string, object>)cogoPointCollected[cornerRecordFormKey])
                                   ["Geolocation_Longitude_s"];
                                var cogoFinalLat = ((Dictionary<string, object>)cogoPointCollected[cornerRecordFormKey])
                                    ["Geolocation_Latitude_s"];

                                cogoFinal.Add("Geolocation_Longitude_s", cogoFinalLong.ToString());
                                cogoFinal.Add("Geolocation_Latitude_s", cogoFinalLat.ToString());
                            }
                        }
                        cr_dictReport.Add("CogoPoint_check", extraCogoPoint_result);
                        cr_dictReport.Add("Layout_check", extraLayout_result);
                        cr_dictReport.Add("Form_Check", listOfFormChecks);
                        cr_dictReport.Add("Form_Result", cornerRecordForms);

                        using (var writer = File.CreateText("CornerRecordForms.json"))
                        {
                            string strResultJson = JsonConvert.SerializeObject(cr_dictReport,
                                Formatting.Indented);
                            writer.WriteLine(strResultJson);
                        }
                    }
                    else
                    {
                        cr_dictReport.Add("Layout_check", extraLayout_result);
                        cr_dictReport.Add("Form_Check", listOfFormChecks);
                        //cr_dictReport.Add("Form_Result", cornerRecordForms);
                        using (var writer = File.CreateText("CornerRecordForms.json"))
                        {
                            string strResultJson = JsonConvert.SerializeObject(cr_dictReport,
                                Formatting.Indented);
                            writer.WriteLine(strResultJson);
                        }
                    }
                    trans.Commit();
                }
            }

            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(("Exception: " + ex.Message));
            }
        }
    }
    #endregion
}