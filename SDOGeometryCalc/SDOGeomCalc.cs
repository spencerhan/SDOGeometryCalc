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
            UID editorUID = new UID();
            editorUID.Value = "esriEditor.Editor";
            IEditor3 editor = (IEditor3)ArcMap.Application.FindExtensionByCLSID(editorUID);
            Boolean editSession = (editor.EditState == esriEditState.esriStateEditing);
            if (!editSession) {
                #region getting active layer
                try {
                    IGeoFeatureLayer iGeoFLayer = this.featureTest(iLayer);
                    IFeatureClass iFclass = iGeoFLayer.FeatureClass;
                    //MessageBox.Show(iFclass.AliasName, "DataSource", MessageBoxButtons.OK);
                    string dbName = string.Concat(iFclass.AliasName.TakeWhile((c) => c != '.')); // reference: https://stackoverflow.com/questions/1857513/get-substring-everything-before-certain-char
                    string layerName = iFclass.AliasName.Replace(dbName, "");
                    string colName;
                    string sdoFunc;
                    esriGeometryType fGeom = iFclass.ShapeType;
                    if (fGeom == esriGeometryType.esriGeometryLine || fGeom == esriGeometryType.esriGeometryPolyline) {
                        colName = "SHAPE_LENGTH";
                        sdoFunc = "SDO_LENGTH";
                        return;
                    } else if (fGeom == esriGeometryType.esriGeometryPolygon) {
                        colName = "SHAPE_AREA";
                        sdoFunc = "SDO_AREA";
                        return;
                    } else {
                        throw new Exception("Not a supported geometry type.");
                    }
                    this.getOracleGeometry(dbName, colName, sdoFunc, layerName, iGeoFLayer);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                }
                #endregion

            } else {
                #region normal labeling
                IGeoFeatureLayer iGeoFLayer = this.featureTest(iLayer);
                this.geometryAnnotation(iGeoFLayer);
                #endregion
            }

        }
        protected override void OnUpdate() {
            Enabled = ArcMap.Application != null;
        }

        protected IGeoFeatureLayer featureTest(ILayer _iLayer) {
            if (_iLayer is IGeoFeatureLayer) {
                IGeoFeatureLayer _iGeoFLayer = (IGeoFeatureLayer)_iLayer;
                return _iGeoFLayer;
            } else {
                throw new InvalidCastException("Please select a layer and check the selected layer is a 'Feature' layer (not a group layer etc.)");
            }
        }

        protected void geometryAnnotation(IGeoFeatureLayer _iGeoFLayer) {
            esriGeometryType _fGeom = _iGeoFLayer.FeatureClass.ShapeType;
            if (_iGeoFLayer.DisplayAnnotation) {
                if (_fGeom == esriGeometryType.esriGeometryLine || _fGeom == esriGeometryType.esriGeometryPolyline) {
                    _iGeoFLayer.DisplayField = "GEOMETRY.LEN";
                } else {
                    _iGeoFLayer.DisplayField = "GEOMETRY.AREA";
                }
            } else {
                if (_fGeom == esriGeometryType.esriGeometryLine || _fGeom == esriGeometryType.esriGeometryPolyline) {
                    _iGeoFLayer.DisplayField = "GEOMETRY.LEN";
                } else {
                    _iGeoFLayer.DisplayField = "GEOMETRY.AREA";
                }
                _iGeoFLayer.DisplayAnnotation = true;
            }
            _iGeoFLayer.ShowTips = true;
        }

        protected void getOracleGeometry(string _dbName, string _colName, string _sdoFunc, string _layerName, IGeoFeatureLayer _iGeoFLayer) {
            string _dbPass;
            if (_colName.Equals("")) {
                throw new Exception("Geometry column name is not found. \n Check layer geometry type (Multiline, line and polygon");
            } else {
                if (_dbName.Equals("GIS_WORKSPACE")) {
                    _dbPass = "";
                    using (OleDbConnection conn = new OleDbConnection()) {
                        conn.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source =(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = oradblive)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = live.wairc.govt.nz))); User ID = " + _dbName + "; Password =" + _dbPass + ";Min Pool Size=10;Connection Lifetime=120;Connection Timeout=60;Incr Pool Size=5; Decr Pool Size=2;Max Pool Size=30;Validate Connection = true";
                        using (OleDbCommand command = conn.CreateCommand()) {
                            command.CommandText = "UPDATE TABLE " + _layerName + " SET " + _colName + "= SDO_GEOM." + _sdoFunc + "(geometry,0.005)";
                            conn.Open();
                            command.ExecuteNonQuery();
                            if (_iGeoFLayer.DisplayAnnotation) {
                                _iGeoFLayer.DisplayField = _colName;
                            } else {
                                _iGeoFLayer.DisplayField = _colName;
                                _iGeoFLayer.DisplayAnnotation = true;
                            }
                            _iGeoFLayer.ShowTips = true;


                            //command.CommandText = "Select" + colName + " from " + layerName + "_EVW"; // feature name needs to be dynamic
                            //using (OleDbDataReader reader = command.ExecuteReader()) {
                            //    List<string> result = new List<string>();
                            //    while (reader.Read()) {
                            //        for (int i = 0; i < reader.FieldCount; i++) {
                            //            result.Add(reader.GetValue(i).ToString());
                            //        }
                            //    }
                            //}

                        }
                    }
                } else {
                    throw new Exception("This is not an Enterprise Geodatabase Layer, please use standard labeling tool.");
                }
            }
        }

    }

}
