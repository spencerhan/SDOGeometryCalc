using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace SDOGeometryCalc {
    public class SDOGeomCalc : ESRI.ArcGIS.Desktop.AddIns.Button {
        public SDOGeomCalc() {
        }

        protected override void OnClick() {
            //
            //  TODO: Sample code showing how to access button host
            //
            ArcMap.Application.CurrentTool = null;

            using (OleDbConnection conn = new OleDbConnection()) {
                string pass = '';
                conn.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source =(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = oradbtest)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = test.wairc.govt.nz))); User ID = gis_workspace; Password ="+ pass +";Min Pool Size=10;Connection Lifetime=120;Connection Timeout=60;Incr Pool Size=5; Decr Pool Size=2;Max Pool Size=30;Validate Connection = true"; 
                OleDbCommand command = conn.CreateCommand();
                command.CommandText = "Select ROUND(SDO_GEOM.SDO_AREA(geometry, 0.005)/10000, 2) AS area_ha from CMO_OTHER_FEATURES_POLYGON_EVW"; // feature name needs to be dynamic
                conn.Open();
                OleDbDataReader reader = command.ExcuteReader();
                List<string> result = new List<string>();
                while (reader.read()){
                    for (int i = 0; i <reader.FieldCount(); i++ ){
                        result.Add(reader.Item[i]);
                    }
                }
              MessageBox.Show(result.ToString(), "Dumping", MessageBoxButtons.OK);             
            }

        }
        protected override void OnUpdate() {
            Enabled = ArcMap.Application != null;
        }
    }

}
