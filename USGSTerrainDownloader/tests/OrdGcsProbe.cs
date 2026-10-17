// Disposable ORD diagnostic. Writes a TEMP report; never writes the DGN or configuration.
using System;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using Bentley.GeoCoordinatesNET;
using Bentley.MstnPlatformNET;

[AddIn(MdlTaskID = "USGSGCSPROBE")]
public sealed class OrdGcsProbe : AddIn
{
    private OrdGcsProbe(IntPtr descriptor) : base(descriptor) { }

    protected override int Run(string[] commandLine)
    {
        string path = Path.Combine(Path.GetTempPath(), "USGSTerrainDownloader-GcsProbe.json");
        var serializer = new JavaScriptSerializer();
        File.WriteAllText(path, "{\"status\":\"running\"}", Encoding.UTF8);
        try
        {
            var model = Session.Instance.GetActiveDgnModelRef();
            if (model == null) throw new InvalidOperationException("No active model.");
            var gcs = DgnGCS.FromModel(model, true);
            if (gcs == null) throw new InvalidOperationException("No primary GCS.");
            string wkt = "", json = "";
            int wktStatus = gcs.GetWellKnownText(out wkt, BaseGCS.WellKnownTextFlavor.wktFlavorEPSG, false);
            int jsonStatus = gcs.ToJson(out json);
            if (String.IsNullOrWhiteSpace(wkt) || String.IsNullOrWhiteSpace(json))
                throw new InvalidOperationException("GCS export returned an empty definition.");
            File.WriteAllText(path, serializer.Serialize(new {
                status = "captured", name = gcs.Name, units = gcs.Units,
                epsg = gcs.GetEPSGCode(false), paper_scale = gcs.PaperScale,
                local_transform_type = gcs.LocalTransformType,
                wkt_status = wktStatus, json_status = jsonStatus, wkt = wkt, gcs_json = json,
                coordinate_agreement_tested = false
            }), Encoding.UTF8);
            return 0;
        }
        catch (Exception error)
        {
            File.WriteAllText(path, serializer.Serialize(new {
                status = "failed", error = error.ToString()
            }), Encoding.UTF8);
            return 1;
        }
    }
}
