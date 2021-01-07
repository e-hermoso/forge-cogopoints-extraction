using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using CivilDB = Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace civil3dCogoPoints
{
    class CogoPointJson
    {
        public static Dictionary<string, object> geolocationCapture(CogoPointCollection passedCogoCollection)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var acDB = doc.Database;
            var civilDoc = CivilApplication.ActiveDocument;

            // Creat a json file of the Corner Record Points found in the CogoPointCollection
            // Retrieves the Long/Lat of the Point and converts to decimal degrees
            // Confirms that the Name field is filled correctly (cr x)

            using (var trans = acDB.TransactionManager.StartTransaction())
            {
                // count the number of duplicte keys "cr#"
                //int dublicate_number = 1;

                Dictionary<string, object> cogoPointJson = new Dictionary<string, object>();
                //Dictionary<string, object> duplicateCogoPoint = new Dictionary<string, object>();

                foreach (ObjectId cogoPointRecord in passedCogoCollection)
                {
                    CogoPoint cogoPointItem = trans.GetObject(cogoPointRecord, OpenMode.ForRead) as CogoPoint;

                    // Access user defined classification in cogopoints table
                    foreach (UDPClassification udpClassification in civilDoc.PointUDPClassifications)
                    {
                        if (udpClassification.Name.ToString() == "Corner Record No.")
                        {
                            //ed.WriteMessage("\n\nUDP Classification: {0}\n", udpClassification.Name);
                            foreach (UDP udp in udpClassification.UDPs)
                            {
                                if (udp.Name.ToString() == "Corner Record #")
                                {
                                    var udpType = udp.GetType().Name;
                                    if (udpType == "UDPString")
                                    {
                                        UDPString udpString = (UDPString)udp;
                                        var cornerRecordNumber = cogoPointItem.GetUDPValue(udpString);
                                        Match crNumMatch = Regex.Match(cornerRecordNumber, "^(\\s*cr\\s*\\d\\d*)$", RegexOptions.IgnoreCase);

                                        if (crNumMatch.Success)
                                        {
                                            //string crNumChecked = cornerRecordNumber.Trim().ToString().ToLower().Replace(" ", "");

                                            Dictionary<String, object> cogoPointGeolocation = new Dictionary<string, object>();

                                            //convert the Lat/Long from Radians to Decimal Degrees
                                            double rad2DegLong = (cogoPointItem.Longitude * 180) / Math.PI;
                                            double rad2DegLat = (cogoPointItem.Latitude * 180) / Math.PI;

                                            //cogoPointGeolocation.Add("Corner_Type_c", "Other");
                                            cogoPointGeolocation.Add("Geolocation_Longitude_s", rad2DegLong);
                                            cogoPointGeolocation.Add("Geolocation_Latitude_s", rad2DegLat);
                                            cogoPointGeolocation.Add("Full Description", cogoPointItem.FullDescription);

                                            // Check if there is a cogoPoint name that already exist.
                                            if (!cogoPointJson.ContainsKey(cornerRecordNumber.Trim().ToString().ToLower().Replace(" ", "")))
                                            {
                                                cogoPointJson.Add(cornerRecordNumber.Trim().ToString().ToLower().Replace(" ", ""), cogoPointGeolocation);
                                            }
                                            else
                                            {
                                                // Issue occurs when a field is named as "cr 1" and "cr1"
                                                //duplicateCogoPoint.Add("Duplicate_Error_" + dublicate_number.ToString(), "Found duplicate corner record name in cogoPoint table: " + cogoPointItem.PointName.Trim().ToString().ToLower().Replace(" ", ""));

                                                //dublicate_number += 1;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                //using (var writer = File.CreateText("CogoPointsGeolocation.json"))
                //{
                //    string strResultJson = JsonConvert.SerializeObject(cogoPointJson,
                //        Formatting.Indented);
                //    writer.WriteLine(strResultJson);
                //}
                trans.Commit();

                // Return cogo point data
                return cogoPointJson;
            }
        }
    }
}
