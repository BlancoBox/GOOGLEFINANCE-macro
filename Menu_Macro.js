function Menu(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("ShepherdMacros");
  menu.addItem("Update", "Take2");
  menu.addItem("Delete", "deletes");
  menu.addToUi();
}
function onOpen(){
Menu();
}
function deletes() {
  var App = SpreadsheetApp;
  var ActiveSheet = App.getActiveSpreadsheet().getActiveSheet();
  
  ActiveSheet.getRange("F2:M18").setValue(null).setFontSize(10);

}
function Take2() {
  var App = SpreadsheetApp;
  var ActiveSheet = App.getActiveSpreadsheet().getActiveSheet();
   
  var Pull = '"Price"';
  var high52 = '"high52"';
  var low52 = '"low52"';
  var closeyest = '"closeyest"';
  var volumeavg ='"volumeavg"';
  var volume = '"volume"';
  var beta = '"beta"';
  var app = ActiveSheet.getRange("A2").getValue();
  var app1 = ActiveSheet.getRange("A3").getValue();
  var app2 = ActiveSheet.getRange("A4").getValue();
  var app3 = ActiveSheet.getRange("A5").getValue();
  var app4 = ActiveSheet.getRange("A6").getValue();
  var app5 = ActiveSheet.getRange("A7").getValue();
  var app6 = ActiveSheet.getRange("A8").getValue();
  var app7 = ActiveSheet.getRange("A9").getValue();
  var app8 = ActiveSheet.getRange("A10").getValue();
  var app9 = ActiveSheet.getRange("A11").getValue();
  var app10 = ActiveSheet.getRange("A12").getValue();
  var app11 = ActiveSheet.getRange("A13").getValue();
  var app12 = ActiveSheet.getRange("A14").getValue();
  var app13 = ActiveSheet.getRange("A15").getValue();
  var app14 = ActiveSheet.getRange("A16").getValue();
  var app15 = ActiveSheet.getRange("A17").getValue();
  var app16 = ActiveSheet.getRange("A18").getValue();
  var s = '"';
  var totav = '=(K2/J2)*100';
  var totav1 = '=(K3/J3)*100';
  var totav2 = '=(K4/J4)*100';
  var totav3 = '=(K5/J5)*100';
  var totav4 = '=(K6/J6)*100';
  var totav5 = '=(K7/J7)*100';
  var totav6 = '=(K8/J8)*100';
  var totav7 = '=(K9/J9)*100';
  var totav8 = '=(K10/J10)*100';
  var totav9 = '=(K11/J11)*100';
  var totav10 = '=(K12/J12)*100';
  var totav11 = '=(K13/J13)*100';
  var totav12 = '=(K14/J14)*100';
  var totav13 = '=(K15/J15)*100';
  var totav14 = '=(K16/J16)*100';
  var totav15 = '=(K17/J17)*100';
  var totav16 = '=(K18/J18)*100';

  var fin = "=GOOGLEFINANCE("+s+app+s+","+Pull+")";
  var fin2 = "=GOOGLEFINANCE("+s+app+s+","+high52+")";
  var fin3 = "=GOOGLEFINANCE("+s+app+s+","+low52+")";
  var fin4 = "=GOOGLEFINANCE("+s+app+s+","+closeyest+")";
  var fin5 = "=GOOGLEFINANCE("+s+app+s+","+volumeavg+")";
  var fin6 = "=GOOGLEFINANCE("+s+app+s+","+volume+")";
  var fin7 = "=GOOGLEFINANCE("+s+app+s+","+beta+")";

  
  var fin_1 = "=GOOGLEFINANCE("+s+app1+s+","+Pull+")";
  var fin2_1 = "=GOOGLEFINANCE("+s+app1+s+","+high52+")";
  var fin3_1 = "=GOOGLEFINANCE("+s+app1+s+","+low52+")";
  var fin4_1 = "=GOOGLEFINANCE("+s+app1+s+","+closeyest+")";
  var fin5_1 = "=GOOGLEFINANCE("+s+app1+s+","+volumeavg+")";
  var fin6_1 = "=GOOGLEFINANCE("+s+app1+s+","+volume+")";
  var fin7_1 = "=GOOGLEFINANCE("+s+app1+s+","+beta+")";
  
  var fin_2 = "=GOOGLEFINANCE("+s+app2+s+","+Pull+")";
  var fin2_2 = "=GOOGLEFINANCE("+s+app2+s+","+high52+")";
  var fin3_2 = "=GOOGLEFINANCE("+s+app2+s+","+low52+")";
  var fin4_2 = "=GOOGLEFINANCE("+s+app2+s+","+closeyest+")";
  var fin5_2 = "=GOOGLEFINANCE("+s+app2+s+","+volumeavg+")";
  var fin6_2 = "=GOOGLEFINANCE("+s+app2+s+","+volume+")";
  var fin7_2 = "=GOOGLEFINANCE("+s+app2+s+","+beta+")";
  
  var fin_3 = "=GOOGLEFINANCE("+s+app3+s+","+Pull+")";
  var fin2_3 = "=GOOGLEFINANCE("+s+app3+s+","+high52+")";
  var fin3_3 = "=GOOGLEFINANCE("+s+app3+s+","+low52+")";
  var fin4_3 = "=GOOGLEFINANCE("+s+app3+s+","+closeyest+")";
  var fin5_3 = "=GOOGLEFINANCE("+s+app3+s+","+volumeavg+")";
  var fin6_3 = "=GOOGLEFINANCE("+s+app3+s+","+volume+")";
  var fin7_3 = "=GOOGLEFINANCE("+s+app3+s+","+beta+")";
  
  var fin_4 = "=GOOGLEFINANCE("+s+app4+s+","+Pull+")";
  var fin2_4 = "=GOOGLEFINANCE("+s+app4+s+","+high52+")";
  var fin3_4 = "=GOOGLEFINANCE("+s+app4+s+","+low52+")";
  var fin4_4 = "=GOOGLEFINANCE("+s+app4+s+","+closeyest+")";
  var fin5_4 = "=GOOGLEFINANCE("+s+app4+s+","+volumeavg+")";
  var fin6_4 = "=GOOGLEFINANCE("+s+app4+s+","+volume+")";
  var fin7_4 = "=GOOGLEFINANCE("+s+app4+s+","+beta+")";
  
  var fin_5 = "=GOOGLEFINANCE("+s+app5+s+","+Pull+")";
  var fin2_5 = "=GOOGLEFINANCE("+s+app5+s+","+high52+")";
  var fin3_5 = "=GOOGLEFINANCE("+s+app5+s+","+low52+")";
  var fin4_5 = "=GOOGLEFINANCE("+s+app5+s+","+closeyest+")";
  var fin5_5 = "=GOOGLEFINANCE("+s+app5+s+","+volumeavg+")";
  var fin6_5 = "=GOOGLEFINANCE("+s+app5+s+","+volume+")";
  var fin7_5 = "=GOOGLEFINANCE("+s+app5+s+","+beta+")";

  var fin_6 = "=GOOGLEFINANCE("+s+app6+s+","+Pull+")";
  var fin2_6 = "=GOOGLEFINANCE("+s+app6+s+","+high52+")";
  var fin3_6 = "=GOOGLEFINANCE("+s+app6+s+","+low52+")";
  var fin4_6 = "=GOOGLEFINANCE("+s+app6+s+","+closeyest+")";
  var fin5_6 = "=GOOGLEFINANCE("+s+app6+s+","+volumeavg+")";
  var fin6_6 = "=GOOGLEFINANCE("+s+app6+s+","+volume+")";
  var fin7_6 = "=GOOGLEFINANCE("+s+app6+s+","+beta+")";

  var fin_7 = "=GOOGLEFINANCE("+s+app7+s+","+Pull+")";
  var fin2_7 = "=GOOGLEFINANCE("+s+app7+s+","+high52+")";
  var fin3_7 = "=GOOGLEFINANCE("+s+app7+s+","+low52+")";
  var fin4_7 = "=GOOGLEFINANCE("+s+app7+s+","+closeyest+")";
  var fin5_7 = "=GOOGLEFINANCE("+s+app7+s+","+volumeavg+")";
  var fin6_7 = "=GOOGLEFINANCE("+s+app7+s+","+volume+")";
  var fin7_7 = "=GOOGLEFINANCE("+s+app7+s+","+beta+")";

  var fin_8 = "=GOOGLEFINANCE("+s+app8+s+","+Pull+")";
  var fin2_8 = "=GOOGLEFINANCE("+s+app8+s+","+high52+")";
  var fin3_8 = "=GOOGLEFINANCE("+s+app8+s+","+low52+")";
  var fin4_8 = "=GOOGLEFINANCE("+s+app8+s+","+closeyest+")";
  var fin5_8 = "=GOOGLEFINANCE("+s+app8+s+","+volumeavg+")";
  var fin6_8 = "=GOOGLEFINANCE("+s+app8+s+","+volume+")";
  var fin7_8 = "=GOOGLEFINANCE("+s+app8+s+","+beta+")";

  var fin_9 = "=GOOGLEFINANCE("+s+app9+s+","+Pull+")";
  var fin2_9 = "=GOOGLEFINANCE("+s+app9+s+","+high52+")";
  var fin3_9 = "=GOOGLEFINANCE("+s+app9+s+","+low52+")";
  var fin4_9 = "=GOOGLEFINANCE("+s+app9+s+","+closeyest+")";
  var fin5_9 = "=GOOGLEFINANCE("+s+app9+s+","+volumeavg+")";
  var fin6_9 = "=GOOGLEFINANCE("+s+app9+s+","+volume+")";
  var fin7_9 = "=GOOGLEFINANCE("+s+app9+s+","+beta+")";

  var fin_10 = "=GOOGLEFINANCE("+s+app10+s+","+Pull+")";
  var fin2_10 = "=GOOGLEFINANCE("+s+app10+s+","+high52+")";
  var fin3_10 = "=GOOGLEFINANCE("+s+app10+s+","+low52+")";
  var fin4_10 = "=GOOGLEFINANCE("+s+app10+s+","+closeyest+")";
  var fin5_10 = "=GOOGLEFINANCE("+s+app10+s+","+volumeavg+")";
  var fin6_10 = "=GOOGLEFINANCE("+s+app10+s+","+volume+")";
  var fin7_10 = "=GOOGLEFINANCE("+s+app10+s+","+beta+")";

  var fin_11 = "=GOOGLEFINANCE("+s+app11+s+","+Pull+")";
  var fin2_11 = "=GOOGLEFINANCE("+s+app11+s+","+high52+")";
  var fin3_11 = "=GOOGLEFINANCE("+s+app11+s+","+low52+")";
  var fin4_11 = "=GOOGLEFINANCE("+s+app11+s+","+closeyest+")";
  var fin5_11 = "=GOOGLEFINANCE("+s+app11+s+","+volumeavg+")";
  var fin6_11 = "=GOOGLEFINANCE("+s+app11+s+","+volume+")";
  var fin7_11 = "=GOOGLEFINANCE("+s+app11+s+","+beta+")";

  var fin_12 = "=GOOGLEFINANCE("+s+app12+s+","+Pull+")";
  var fin2_12 = "=GOOGLEFINANCE("+s+app12+s+","+high52+")";
  var fin3_12 = "=GOOGLEFINANCE("+s+app12+s+","+low52+")";
  var fin4_12 = "=GOOGLEFINANCE("+s+app12+s+","+closeyest+")";
  var fin5_12 = "=GOOGLEFINANCE("+s+app12+s+","+volumeavg+")";
  var fin6_12 = "=GOOGLEFINANCE("+s+app12+s+","+volume+")";
  var fin7_12 = "=GOOGLEFINANCE("+s+app12+s+","+beta+")";

  var fin_13 = "=GOOGLEFINANCE("+s+app13+s+","+Pull+")";
  var fin2_13 = "=GOOGLEFINANCE("+s+app13+s+","+high52+")";
  var fin3_13 = "=GOOGLEFINANCE("+s+app13+s+","+low52+")";
  var fin4_13 = "=GOOGLEFINANCE("+s+app13+s+","+closeyest+")";
  var fin5_13 = "=GOOGLEFINANCE("+s+app13+s+","+volumeavg+")";
  var fin6_13 = "=GOOGLEFINANCE("+s+app13+s+","+volume+")";
  var fin7_13 = "=GOOGLEFINANCE("+s+app13+s+","+beta+")";

  var fin_14 = "=GOOGLEFINANCE("+s+app14+s+","+Pull+")";
  var fin2_14 = "=GOOGLEFINANCE("+s+app14+s+","+high52+")";
  var fin3_14 = "=GOOGLEFINANCE("+s+app14+s+","+low52+")";
  var fin4_14 = "=GOOGLEFINANCE("+s+app14+s+","+closeyest+")";
  var fin5_14 = "=GOOGLEFINANCE("+s+app14+s+","+volumeavg+")";
  var fin6_14 = "=GOOGLEFINANCE("+s+app14+s+","+volume+")";
  var fin7_14 = "=GOOGLEFINANCE("+s+app14+s+","+beta+")";

  var fin_15 = "=GOOGLEFINANCE("+s+app15+s+","+Pull+")";
  var fin2_15 = "=GOOGLEFINANCE("+s+app15+s+","+high52+")";
  var fin3_15 = "=GOOGLEFINANCE("+s+app15+s+","+low52+")";
  var fin4_15 = "=GOOGLEFINANCE("+s+app15+s+","+closeyest+")";
  var fin5_15 = "=GOOGLEFINANCE("+s+app15+s+","+volumeavg+")";
  var fin6_15 = "=GOOGLEFINANCE("+s+app15+s+","+volume+")";
  var fin7_15 = "=GOOGLEFINANCE("+s+app15+s+","+beta+")";

  var fin_16 = "=GOOGLEFINANCE("+s+app16+s+","+Pull+")";
  var fin2_16 = "=GOOGLEFINANCE("+s+app16+s+","+high52+")";
  var fin3_16 = "=GOOGLEFINANCE("+s+app16+s+","+low52+")";
  var fin4_16 = "=GOOGLEFINANCE("+s+app16+s+","+closeyest+")";
  var fin5_16 = "=GOOGLEFINANCE("+s+app16+s+","+volumeavg+")";
  var fin6_16 = "=GOOGLEFINANCE("+s+app16+s+","+volume+")";
  var fin7_16 = "=GOOGLEFINANCE("+s+app16+s+","+beta+")";
  
  SpreadsheetApp.flush();
  
 if (app != null) {
  ActiveSheet.getRange("F2").setValue(fin).setFontSize(10);
  ActiveSheet.getRange("G2").setValue(fin2).setFontSize(10);
  ActiveSheet.getRange("H2").setValue(fin3).setFontSize(10);
  ActiveSheet.getRange("I2").setValue(fin4).setFontSize(10);
  ActiveSheet.getRange("J2").setValue(fin5).setFontSize(10);
  ActiveSheet.getRange("K2").setValue(fin6).setFontSize(10);
  ActiveSheet.getRange("L2").setValue(totav).setFontSize(10);
  ActiveSheet.getRange("M2").setValue(fin7).setFontSize(10);
 }else{  ActiveSheet.getRange("F2:M2").setValue(null).setFontSize(10);}
  
   if (app1 != null) {
  ActiveSheet.getRange("F3").setValue(fin_1).setFontSize(10);
  ActiveSheet.getRange("G3").setValue(fin2_1).setFontSize(10);
  ActiveSheet.getRange("H3").setValue(fin3_1).setFontSize(10);
  ActiveSheet.getRange("I3").setValue(fin4_1).setFontSize(10);
  ActiveSheet.getRange("J3").setValue(fin5_1).setFontSize(10);
  ActiveSheet.getRange("K3").setValue(fin6_1).setFontSize(10);
  ActiveSheet.getRange("L3").setValue(totav1).setFontSize(10);
  ActiveSheet.getRange("M3").setValue(fin7_1).setFontSize(10);
   }else{  ActiveSheet.getRange("F3:M3").setValue(null).setFontSize(10);}
  
   if (app2 != null) {
  ActiveSheet.getRange("F4").setValue(fin_2).setFontSize(10);
  ActiveSheet.getRange("G4").setValue(fin2_2).setFontSize(10);
  ActiveSheet.getRange("H4").setValue(fin3_2).setFontSize(10);
  ActiveSheet.getRange("I4").setValue(fin4_2).setFontSize(10);
  ActiveSheet.getRange("J4").setValue(fin5_2).setFontSize(10);
  ActiveSheet.getRange("K4").setValue(fin6_2).setFontSize(10);
  ActiveSheet.getRange("L4").setValue(totav2).setFontSize(10);
  ActiveSheet.getRange("M4").setValue(fin7_2).setFontSize(10);
   }else{  ActiveSheet.getRange("F4:M4").setValue(null).setFontSize(10);}
  
     if (app3 != null) {
  ActiveSheet.getRange("F5").setValue(fin_3).setFontSize(10);
  ActiveSheet.getRange("G5").setValue(fin2_3).setFontSize(10);
  ActiveSheet.getRange("H5").setValue(fin3_3).setFontSize(10);
  ActiveSheet.getRange("I5").setValue(fin4_3).setFontSize(10);
  ActiveSheet.getRange("J5").setValue(fin5_3).setFontSize(10);
  ActiveSheet.getRange("K5").setValue(fin6_3).setFontSize(10);
  ActiveSheet.getRange("L5").setValue(totav3).setFontSize(10);
  ActiveSheet.getRange("M5").setValue(fin7_3).setFontSize(10);
     }else{  ActiveSheet.getRange("F5:M5").setValue(null).setFontSize(10);}
  
     if (app4 != null) {
  ActiveSheet.getRange("F6").setValue(fin_4).setFontSize(10);
  ActiveSheet.getRange("G6").setValue(fin2_4).setFontSize(10);
  ActiveSheet.getRange("H6").setValue(fin3_4).setFontSize(10);
  ActiveSheet.getRange("I6").setValue(fin4_4).setFontSize(10);
  ActiveSheet.getRange("J6").setValue(fin5_4).setFontSize(10);
  ActiveSheet.getRange("K6").setValue(fin6_4).setFontSize(10);
  ActiveSheet.getRange("L6").setValue(totav4).setFontSize(10);
  ActiveSheet.getRange("M6").setValue(fin7_4).setFontSize(10);
     }else{  ActiveSheet.getRange("F6:M6").setValue(null).setFontSize(10);}
  
     if (app5 != null) {
  ActiveSheet.getRange("F7").setValue(fin_5).setFontSize(10);
  ActiveSheet.getRange("G7").setValue(fin2_5).setFontSize(10);
  ActiveSheet.getRange("H7").setValue(fin3_5).setFontSize(10);
  ActiveSheet.getRange("I7").setValue(fin4_5).setFontSize(10);
  ActiveSheet.getRange("J7").setValue(fin5_5).setFontSize(10);
  ActiveSheet.getRange("K7").setValue(fin6_5).setFontSize(10);
  ActiveSheet.getRange("L7").setValue(totav5).setFontSize(10);
  ActiveSheet.getRange("M7").setValue(fin7_5).setFontSize(10);
     }else{  ActiveSheet.getRange("F7:M7").setValue(null).setFontSize(10);}
      
     if (app6 != null) {
  ActiveSheet.getRange("F8").setValue(fin_6).setFontSize(10);
  ActiveSheet.getRange("G8").setValue(fin2_6).setFontSize(10);
  ActiveSheet.getRange("H8").setValue(fin3_6).setFontSize(10);
  ActiveSheet.getRange("I8").setValue(fin4_6).setFontSize(10);
  ActiveSheet.getRange("J8").setValue(fin5_6).setFontSize(10);
  ActiveSheet.getRange("K8").setValue(fin6_6).setFontSize(10);
  ActiveSheet.getRange("L8").setValue(totav6).setFontSize(10);
  ActiveSheet.getRange("M8").setValue(fin7_6).setFontSize(10);
     }else{  ActiveSheet.getRange("F8:M8").setValue(null).setFontSize(10);}
  
     if (app7 != null) {
  ActiveSheet.getRange("F9").setValue(fin_7).setFontSize(10);
  ActiveSheet.getRange("G9").setValue(fin2_7).setFontSize(10);
  ActiveSheet.getRange("H9").setValue(fin3_7).setFontSize(10);
  ActiveSheet.getRange("I9").setValue(fin4_7).setFontSize(10);
  ActiveSheet.getRange("J9").setValue(fin5_7).setFontSize(10);
  ActiveSheet.getRange("K9").setValue(fin6_7).setFontSize(10);
  ActiveSheet.getRange("L9").setValue(totav7).setFontSize(10);
  ActiveSheet.getRange("M9").setValue(fin7_7).setFontSize(10);
     }else{  ActiveSheet.getRange("F8:M8").setValue(null).setFontSize(10);}
  
       if (app8 != null) {
  ActiveSheet.getRange("F10").setValue(fin_8).setFontSize(10);
  ActiveSheet.getRange("G10").setValue(fin2_8).setFontSize(10);
  ActiveSheet.getRange("H10").setValue(fin3_8).setFontSize(10);
  ActiveSheet.getRange("I10").setValue(fin4_8).setFontSize(10);
  ActiveSheet.getRange("J10").setValue(fin5_8).setFontSize(10);
  ActiveSheet.getRange("K10").setValue(fin6_8).setFontSize(10);
  ActiveSheet.getRange("L10").setValue(totav8).setFontSize(10);
  ActiveSheet.getRange("M10").setValue(fin7_8).setFontSize(10);
     }else{  ActiveSheet.getRange("F9:M9").setValue(null).setFontSize(10);}
  
       if (app9 != null) {
  ActiveSheet.getRange("F11").setValue(fin_9).setFontSize(10);
  ActiveSheet.getRange("G11").setValue(fin2_9).setFontSize(10);
  ActiveSheet.getRange("H11").setValue(fin3_9).setFontSize(10);
  ActiveSheet.getRange("I11").setValue(fin4_9).setFontSize(10);
  ActiveSheet.getRange("J11").setValue(fin5_9).setFontSize(10);
  ActiveSheet.getRange("K11").setValue(fin6_9).setFontSize(10);
  ActiveSheet.getRange("L11").setValue(totav9).setFontSize(10);
  ActiveSheet.getRange("M11").setValue(fin7_9).setFontSize(10);
     }else{  ActiveSheet.getRange("F11:M11").setValue(null).setFontSize(10);}
  
       if (app10 != null) {
  ActiveSheet.getRange("F12").setValue(fin_10).setFontSize(10);
  ActiveSheet.getRange("G12").setValue(fin2_10).setFontSize(10);
  ActiveSheet.getRange("H12").setValue(fin3_10).setFontSize(10);
  ActiveSheet.getRange("I12").setValue(fin4_10).setFontSize(10);
  ActiveSheet.getRange("J12").setValue(fin5_10).setFontSize(10);
  ActiveSheet.getRange("K12").setValue(fin6_10).setFontSize(10);
  ActiveSheet.getRange("L12").setValue(totav10).setFontSize(10);
  ActiveSheet.getRange("M12").setValue(fin7_10).setFontSize(10);
     }else{  ActiveSheet.getRange("F12:M12").setValue(null).setFontSize(10);}
  
       if (app11 != null) {
  ActiveSheet.getRange("F13").setValue(fin_11).setFontSize(10);
  ActiveSheet.getRange("G13").setValue(fin2_11).setFontSize(10);
  ActiveSheet.getRange("H13").setValue(fin3_11).setFontSize(10);
  ActiveSheet.getRange("I13").setValue(fin4_11).setFontSize(10);
  ActiveSheet.getRange("J13").setValue(fin5_11).setFontSize(10);
  ActiveSheet.getRange("K13").setValue(fin6_11).setFontSize(10);
  ActiveSheet.getRange("L13").setValue(totav11).setFontSize(10);
  ActiveSheet.getRange("M13").setValue(fin7_11).setFontSize(10);
     }else{  ActiveSheet.getRange("F13:M13").setValue(null).setFontSize(10);}
  
   if (app12 != null) {
  ActiveSheet.getRange("F14").setValue(fin_12).setFontSize(10);
  ActiveSheet.getRange("G14").setValue(fin2_12).setFontSize(10);
  ActiveSheet.getRange("H14").setValue(fin3_12).setFontSize(10);
  ActiveSheet.getRange("I14").setValue(fin4_12).setFontSize(10);
  ActiveSheet.getRange("J14").setValue(fin5_12).setFontSize(10);
  ActiveSheet.getRange("K14").setValue(fin6_12).setFontSize(10);
  ActiveSheet.getRange("L14").setValue(totav12).setFontSize(10);
  ActiveSheet.getRange("M14").setValue(fin7_12).setFontSize(10);
   }else{  ActiveSheet.getRange("F14:M14").setValue(null).setFontSize(10);}
  
   if (app13 != null) {
  ActiveSheet.getRange("F15").setValue(fin_13).setFontSize(10);
  ActiveSheet.getRange("G15").setValue(fin2_13).setFontSize(10);
  ActiveSheet.getRange("H15").setValue(fin3_13).setFontSize(10);
  ActiveSheet.getRange("I15").setValue(fin4_13).setFontSize(10);
  ActiveSheet.getRange("J15").setValue(fin5_13).setFontSize(10);
  ActiveSheet.getRange("K15").setValue(fin6_13).setFontSize(10);
  ActiveSheet.getRange("L15").setValue(totav13).setFontSize(10);
  ActiveSheet.getRange("M15").setValue(fin7_13).setFontSize(10);
   }else{  ActiveSheet.getRange("F15:M15").setValue(null).setFontSize(10);}  
  
   if (app14 != null) {
  ActiveSheet.getRange("F16").setValue(fin_14).setFontSize(10);
  ActiveSheet.getRange("G16").setValue(fin2_14).setFontSize(10);
  ActiveSheet.getRange("H16").setValue(fin3_14).setFontSize(10);
  ActiveSheet.getRange("I16").setValue(fin4_14).setFontSize(10);
  ActiveSheet.getRange("J16").setValue(fin5_14).setFontSize(10);
  ActiveSheet.getRange("K16").setValue(fin6_14).setFontSize(10);
  ActiveSheet.getRange("L16").setValue(totav14).setFontSize(10);
  ActiveSheet.getRange("M16").setValue(fin7_14).setFontSize(10);
   }else{  ActiveSheet.getRange("F16:M16").setValue(null).setFontSize(10);}  
  
  
   if (app15 != null) {
  ActiveSheet.getRange("F17").setValue(fin_15).setFontSize(10);
  ActiveSheet.getRange("G17").setValue(fin2_15).setFontSize(10);
  ActiveSheet.getRange("H17").setValue(fin3_15).setFontSize(10);
  ActiveSheet.getRange("I17").setValue(fin4_15).setFontSize(10);
  ActiveSheet.getRange("J17").setValue(fin5_15).setFontSize(10);
  ActiveSheet.getRange("K17").setValue(fin6_15).setFontSize(10);
  ActiveSheet.getRange("L17").setValue(totav15).setFontSize(10);
  ActiveSheet.getRange("M17").setValue(fin7_15).setFontSize(10);
   }else{  ActiveSheet.getRange("F17:M17").setValue(null).setFontSize(10);}  
  
   if (app16 != null) {
  ActiveSheet.getRange("F18").setValue(fin_16).setFontSize(10);
  ActiveSheet.getRange("G18").setValue(fin2_16).setFontSize(10);
  ActiveSheet.getRange("H18").setValue(fin3_16).setFontSize(10);
  ActiveSheet.getRange("I18").setValue(fin4_16).setFontSize(10);
  ActiveSheet.getRange("J18").setValue(fin5_16).setFontSize(10);
  ActiveSheet.getRange("K18").setValue(fin6_16).setFontSize(10);
  ActiveSheet.getRange("L18").setValue(totav16).setFontSize(10);
  ActiveSheet.getRange("M18").setValue(fin7_16).setFontSize(10);
   }else{  ActiveSheet.getRange("F18:M18").setValue(null).setFontSize(10);}  
  
  SpreadsheetApp.flush();

}
