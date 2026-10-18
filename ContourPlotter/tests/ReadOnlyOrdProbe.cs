// SPDX-License-Identifier: MIT
using System;
using System.Collections.Generic;
using Bentley.Interop.MicroStationDGN;
public static class ContourReadOnlyProbe {
 public static Dictionary<string,object> Read(object instance) {
  var app = (Application)instance; var model = app.ActiveModelReference; var unit = model.get_MasterUnit();
  var result = new Dictionary<string,object>();
  result["unit"] = new Dictionary<string,object>{{"label",unit.Label},{"base",unit.Base.ToString()},
   {"numerator",unit.UnitsPerBaseNumerator},{"denominator",unit.UnitsPerBaseDenominator},
   {"metresPerMasterUnit",unit.UnitsPerBaseDenominator/unit.UnitsPerBaseNumerator}};
  var names = new List<string>();
  foreach(Level level in app.ActiveDesignFile.Levels)
   if(!level.IsFromLevelLibrary && !level.IsLocked && level.ElementAccess==MsdLevelElementAccess.All) names.Add(level.Name);
  result["writableLocalLevelCount"]=names.Count; result["writableLocalLevels"]=names;
  var styles=new List<string>(); TextStyle chosen=null;
  foreach(TextStyle style in app.ActiveDesignFile.TextStyles)
   if(!style.IsFromLibrary && !String.IsNullOrEmpty(style.Name)) {
    styles.Add(style.Name); if(chosen==null && style.Height>0 && style.Width>0) chosen=style;
   }
  result["localTextStyles"]=styles;
  var vertices=new Point3d[1024]; for(int i=0;i<vertices.Length;i++){vertices[i].X=i;vertices[i].Y=i%2;}
  var line=app.CreateLineElement1(null,ref vertices);
  line.Color=app.ByLevelColor(); line.LineWeight=app.ByLevelLineWeight(); line.LineStyle=app.ByLevelLineStyle();
  result["preparedLineID"]=line.ID64.ToString(); result["preparedLineVertices"]=vertices.Length;
  if(chosen!=null) {
   double originalHeight=chosen.Height,originalWidth=chosen.Width; string styleName=chosen.Name;
   Point3d origin=new Point3d(); Matrix3d rotation=app.Matrix3dIdentity();
   var text=app.CreateTextElement1(null,"250.5",ref origin,ref rotation);
   text.TextStyle=chosen;
   text.TextStyle.Height=5; text.TextStyle.Width=5*originalWidth/originalHeight;
   text.TextStyle.Justification=MsdTextJustification.CenterCenter;
   text.TextStyle.UsesColor=false; text.TextStyle.BorderAndBackgroundVisible=false;
   var local=text.TextStyle; text.TextStyle=local;
   double width=0,height=0; text.GetTotalTextSize(out width,out height);
   result["initialMeasuredHeight"]=height;
   text.ScaleUniform(ref origin,5/height);
   text.GetTotalTextSize(out width,out height);
   var again=app.ActiveDesignFile.TextStyles.Find(styleName,null);
   result["preparedText"]=new Dictionary<string,object>{{"style",styleName},{"id",text.ID64.ToString()},
    {"measuredWidth",width},{"measuredHeight",height},{"localHeight",text.TextStyle.Height},
    {"sharedStyleUnchanged",again.Height==originalHeight && again.Width==originalWidth}};
  }
  return result;
 }
}