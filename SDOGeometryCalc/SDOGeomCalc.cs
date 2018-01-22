using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Editor;

namespace SDOGeometryCalc {
    public class SDOGeomCalc : ESRI.ArcGIS.Desktop.AddIns.Button {
        private static Boolean labelOff = true; // check label on or not.
        public SDOGeomCalc() {
        }

        protected override void OnClick() {
            //
            //  TODO: Sample code showing how to access button host
            //
            ArcMap.Application.CurrentTool = null;
            IMxDocument pMxDoc = (IMxDocument)ArcMap.Application.Document;
            IMap pMapWin = pMxDoc.FocusMap;
            ILayer iLayer = pMxDoc.SelectedLayer;

            if (labelOff) {
                #region getting active layer
                try {
                    if (iLayer is IFeatureLayer) {
                        IFeatureLayer iFLayer = (IFeatureLayer)iLayer;
                        IFeatureClass iFclass = iFLayer.FeatureClass;
                        //MessageBox.Show(iFclass.AliasName, "DataSource", MessageBoxButtons.OK);
                        string dbName = string.Concat(iFclass.AliasName.TakeWhile((c) => c != '.')); // reference: https://stackoverflow.com/questions/1857513/get-substring-everything-before-certain-char
                        string layerName = iFclass.AliasName.Replace(dbName, "");
                        string colName;
                        string sdoFunc;
                        UID editorUID = new UID();
                        editorUID.Value = "esriEditor.Editor";
                        IEditor3 editor = (IEditor3)ArcMap.Application.FindExtensionByCLSID(editorUID);
                        Boolean editSession = (editor.EditState == esriEditState.esriStateEditing);
                        esriGeometryType fGeom = iFclass.ShapeType;
                        if (editSession) {
                            throw new Exception("Please exit current Editing Session.");
                        } else {
                            if (fGeom == esriGeometryType.esriGeometryLine || fGeom == esriGeometryType.esriGeometryPolyline) {
                                colName = "SHAPE_LENGTH";
                                sdoFunc = "SDO_LENGTH";
                                return;
                            } else if (fGeom == esriGeometryType.esriGeometryPolygon) {
                                colName = "SHAPE_AREA";
                                sdoFunc = "SDO_AREA";
                                return;
                            } else {
                                colName = "";
                                sdoFunc = "";
                            }
                            this.getOracleGeometry(dbName, colName, sdoFunc, layerName + "_EVW");// tailing of actual layername
                            labelOff = false;
                        }
                    } else {
                        throw new InvalidCastException("Please select a layer and check the selected layer is a 'Feature' layer (not a group layer etc.)");
                    }
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                }
                #endregion


            } else {
                #region removing label

                #endregion
            }

        }
        protected override void OnUpdate() {
            Enabled = ArcMap.Application != null;
        }

        protected void getOracleGeometry(string dbName, string colName, string sdoFunc, string layerName) {
            string dbPass;
            if (colName.Equals("")) {
                throw new Exception("Geometry column name is not found. \n Check layer geometry type (Multiline, line and polygon");
            } else {
                if (dbName.Equals("GIS_WORKSPACE")) {
                    dbPass = "";
                    using (OleDbConnection conn = new OleDbConnection()) {
                        conn.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source =(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = oradblive)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = live.wairc.govt.nz))); User ID = " + dbName + "; Password =" + dbPass + ";Min Pool Size=10;Connection Lifetime=120;Connection Timeout=60;Incr Pool Size=5; Decr Pool Size=2;Max Pool Size=30;Validate Connection = true";
                        using (OleDbCommand command = conn.CreateCommand()) {
                            command.CommandText = "UPDATE TABLE " + layerName + " SET " + colName + "= SDO_GEOM." + sdoFunc + "(geometry,0.005)";
                            conn.Open();


                            #region Reading data
                            command.CommandText = "Select ROUND(SDO_GEOM.SDO_AREA(geometry, 0.005)/10000, 6) AS area_ha from CMO_OTHER_FEATURES_POLYGON_EVW"; // feature name needs to be dynamic
                            using (OleDbDataReader reader = command.ExecuteReader()) {
                                List<string> result = new List<string>();
                                while (reader.Read()) {
                                    for (int i = 0; i < reader.FieldCount; i++) {
                                        result.Add(reader.GetValue(i).ToString());
                                    }
                                }
                            }
                            #endregion
                        }
                    }
                } else {
                    throw new Exception("This is not an Enterprise Geodatabase Layer, please use standard labeling tool.");
                }
            }
        }

    }

}
