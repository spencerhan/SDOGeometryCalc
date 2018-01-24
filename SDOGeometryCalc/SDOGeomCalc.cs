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
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Framework;

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
            IActiveView aView = pMxDoc.ActiveView;
            IMap pMapWin = pMxDoc.FocusMap;
            ILayer iLayer = pMxDoc.SelectedLayer;
            UID editorUID = new UID();
            editorUID.Value = "esriEditor.Editor";
            IEditor3 editor = (IEditor3)ArcMap.Application.FindExtensionByCLSID(editorUID);
            bool editSession = (editor.EditState == esriEditState.esriStateEditing);
            if (!editSession) {
                #region getting active layer
                try {
                    IGeoFeatureLayer iGeoFLayer = this.featureTest(iLayer);
                    IFeatureClass iFclass = iGeoFLayer.FeatureClass;
                    //MessageBox.Show(iFclass.AliasName, "DataSource", MessageBoxButtons.OK);
                    string dbName = string.Concat(iFclass.AliasName.TakeWhile((c) => c != '.')); // reference: https://stackoverflow.com/questions/1857513/get-substring-everything-before-certain-char
                    string layerName = iFclass.AliasName.Replace(dbName + ".", "");
                    string colName;
                    string sdoFunc;
                    esriGeometryType fGeom = iFclass.ShapeType;
                    if (fGeom == esriGeometryType.esriGeometryLine || fGeom == esriGeometryType.esriGeometryPolyline) {
                        colName = "SHAPE_LENGTH";
                        sdoFunc = "SDO_LENGTH";
                    } else if (fGeom == esriGeometryType.esriGeometryPolygon) {
                        colName = "SHAPE_AREA";
                        sdoFunc = "SDO_AREA";
                    } else {
                        throw new Exception("Not a supported geometry type.");
                    }
                    this.getOracleGeometry(dbName, colName, sdoFunc, layerName, iGeoFLayer, aView);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                }
                #endregion

            } else {
                #region normal labeling
                IGeoFeatureLayer iGeoFLayer = this.featureTest(iLayer);
                if (iGeoFLayer != null) {
                    esriGeometryType fGeom = iGeoFLayer.FeatureClass.ShapeType;
                    if (iGeoFLayer.DisplayAnnotation) {
                        if (fGeom == esriGeometryType.esriGeometryLine || fGeom == esriGeometryType.esriGeometryPolyline) {
                            this.labeling(iGeoFLayer, "[GEOMETRY.LEN]", 0, 1000000, false, false, aView);
                        } else {
                            this.labeling(iGeoFLayer, "[GEOMETRY.AREA]", 0, 1000000, false, false, aView);
                        }
                    } else {
                        if (fGeom == esriGeometryType.esriGeometryLine || fGeom == esriGeometryType.esriGeometryPolyline) {
                            this.labeling(iGeoFLayer, "[GEOMETRY.LEN]", 0, 1000000, true, true, aView);
                        } else {
                            this.labeling(iGeoFLayer, "[GEOMETRY.AREA]", 0, 1000000, true, true, aView);
                        }
                    }
                } else {
                    throw new Exception("Feature class is empty (null).");
                }

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

        protected void getOracleGeometry(string _dbName, string _colName, string _sdoFunc, string _layerName, IGeoFeatureLayer _iGeoFLayer, IActiveView _aView) {
            string _dbPass;
            if (_colName.Equals("")) {
                throw new Exception("Geometry column name is not found. \n Check layer geometry type (Multiline, line and polygon");
            } else {
                if (_dbName.Equals("GIS_WORKSPACE")) {
                    _dbPass = "";
                    using (OleDbConnection conn = new OleDbConnection()) {
                        OleDbTransaction transaction = null;
                        conn.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source =(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = oradbtest)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = test.wairc.govt.nz))); User ID =" + _dbName + "; Password =" + _dbPass + ";Min Pool Size=10;Connection Lifetime=120;Connection Timeout=60;Incr Pool Size=5; Decr Pool Size=2;Max Pool Size=30;Validate Connection = true";
                        using (OleDbCommand command = conn.CreateCommand()) {
                            command.CommandText = "UPDATE " + _layerName + " SET " + _colName + "= SDO_GEOM." + _sdoFunc + "(GEOMETRY,0.005) WHERE GEOMETRY is not null";
                            try {
                                conn.Open();
                                transaction = conn.BeginTransaction();
                                command.Transaction = transaction;
                                command.ExecuteNonQuery();
                                transaction.Commit();
                                if (_iGeoFLayer.DisplayAnnotation) {
                                    this.labeling(_iGeoFLayer, "["+_colName+"]", 0, 1000000, false, false, _aView);
                                } else {
                                    this.labeling(_iGeoFLayer, "[" + _colName + "]", 0, 1000000, true, true, _aView);
                                }
                            
                            } catch (Exception ex) {
                                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                                try {
                                    transaction.Rollback();
                                } catch {

                                }
                            }

                        }
                    }
                } else {
                    throw new Exception("This is not an Enterprise Geodatabase Layer, please use standard labeling tool.");
                }
            }
        }
        public void labeling(IGeoFeatureLayer _iGeoFLayer, string _labelFuc, double _maxScale, double _minScale, bool _showMapTips, bool _displayAnnotation, IActiveView _aView) {
            RgbColor _labelColor = new RgbColor();
            _labelColor.RGB = Microsoft.VisualBasic.Information.RGB(0, 0, 0);
            IAnnotateLayerPropertiesCollection propertiesColl = _iGeoFLayer.AnnotationProperties;
            IAnnotateLayerProperties labelEngineProperties = new LabelEngineLayerProperties() as IAnnotateLayerProperties;
            IElementCollection placedElements = new ElementCollection();
            IElementCollection unplacedElements = new ElementCollection();
            propertiesColl.QueryItem(0, out labelEngineProperties, out placedElements, out unplacedElements);
            ILabelEngineLayerProperties lpLabelEngine = labelEngineProperties as ILabelEngineLayerProperties;
            lpLabelEngine.Expression = _labelFuc;
            lpLabelEngine.Symbol.Color = _labelColor;
            labelEngineProperties.AnnotationMaximumScale = _maxScale; // closer
            labelEngineProperties.AnnotationMinimumScale = _minScale; // further
            IFeatureLayer thisFeatureLayer = _iGeoFLayer as IFeatureLayer;
            IDisplayString displayString = thisFeatureLayer as IDisplayString;
            IDisplayExpressionProperties properties = displayString.ExpressionProperties;
            properties.Expression = _labelFuc; //example: "[OWNER_NAME] & vbnewline & \"$\" & [TAX_VALUE]";
            _iGeoFLayer.DisplayAnnotation = _displayAnnotation;
            _iGeoFLayer.ShowTips = _showMapTips;

            // refresh map window 
            _aView.Refresh();
        }
         
    }
}
