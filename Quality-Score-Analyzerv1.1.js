/* Adwords Quality-Score Analyzer
 * Description: Analyze the Quality Score components such as Ad relevance,Exp.CTR and Landing page.
 * Author:RitwikGA
 * Version 1.1
 * Copyright (c) 2016 Licensed under GPL licenses.
 * Mail: ritwikga@gmail.com
 */

function  main()
{
  var AccountName=AdWordsApp.currentAccount().getName()
 
  //Create Spreadsheet 
  var url=getSpreadsheetURL("Rutu-"+AccountName+"_Q_Score");
  var spreadsheet = SpreadsheetApp.openByUrl(url)
  
  //Create Sheets
  var sh2=spreadsheet.getSheets()[0].setName("Q_Score")
  var sh3=getsheet(spreadsheet,"RAW_D");
  
Q_Score(sh2,sh3)
Charts_(sh2)
}
function Q_Score(sh2,sh3) {

var Tzone=AdWordsApp.currentAccount().getTimeZone();  
  sh2.getRange(1,1,1,1).setValue("Start Date ---->")
  sh2.getRange(1,4,1,1).setValue("<-------End Date")
  sh2.getRange("B1:C1").setBackground("#cccccc")
  sh2.getRange(3,1,1,1).setValue("Keyword Filter->")
  sh2.getRange(3,2,1,1).setBackground("#cccccc")
  
  
  ///Fetch Date (if Date is entered)
  var start=typeof(sh2.getRange(1,2,1,1).getValue())=="object"?Utilities.formatDate(sh2.getRange(1,2,1,1).getValue(), Tzone, "yyyyMMdd"):"LAST_7_DAYS";
  var end = typeof(sh2.getRange(1,3,1,1).getValue())=="object"?Utilities.formatDate(sh2.getRange(1,3,1,1).getValue(), Tzone, "yyyyMMdd"):""; 
  
  
  //Default Last 7 Days Data (if Date not Entered)
  
  if(start=="LAST_7_DAYS")
  {
  var report = AdWordsApp.report(
     "SELECT Criteria,CampaignName,Clicks,Impressions,Cost,Ctr,QualityScore,AveragePosition,"+
     "SearchImpressionShare,Conversions,SearchPredictedCtr,CreativeQualityScore,"+
     "PostClickQualityScore,Status,SystemServingStatus "+
     "FROM   KEYWORDS_PERFORMANCE_REPORT WHERE Impressions>0 AND Status = 'ENABLED' AND AdGroupStatus = 'ENABLED' AND CampaignStatus = 'ENABLED' "+ 
     "DURING "+start);
     report.exportToSheet(sh3);
     sh2.getRange(1,6,1,1).setValue("Date Range:Last 7 Days")
  } else
  {
  var report = AdWordsApp.report(
     "SELECT Criteria,CampaignName,Clicks,Impressions,Cost,Ctr,QualityScore,AveragePosition,"+
     "SearchImpressionShare,Conversions,SearchPredictedCtr,CreativeQualityScore,"+
     "PostClickQualityScore,Status,SystemServingStatus "+
     "FROM   KEYWORDS_PERFORMANCE_REPORT WHERE Impressions>0 AND Status = 'ENABLED' AND AdGroupStatus = 'ENABLED' AND CampaignStatus = 'ENABLED' "+ 
     "DURING "+start+","+end);
     report.exportToSheet(sh3);
     sh2.getRange(1,6,1,1).setValue("Date Range:"+start+" "+end)
  }
    
sh3.sort(5,false)

//Query the exported data to extraxt the relevant info
sh3.getRange(2,27,1,1).setValue("=QUERY($A:$O,IF("+sh2.getName()+"!$B$3=\"\",CONCATENATE(\"select L,count(A) group by L\"),CONCATENATE(\"select L,count(A) where A contains'\","+sh2.getName()+"!$B$3,\"' group by L\")),1)")
sh3.getRange(2,29,1,1).setValue("=QUERY($A:$O,IF("+sh2.getName()+"!$B$3=\"\",CONCATENATE(\"select K,count(A) group by K\"),CONCATENATE(\"select K,count(A) where A contains'\","+sh2.getName()+"!$B$3,\"' group by K\")),1)")
sh3.getRange(2,31,1,1).setValue("=QUERY($A:$O,IF("+sh2.getName()+"!$B$3=\"\",CONCATENATE(\"select M,count(A) group by M\"),CONCATENATE(\"select M,count(A) where A contains'\","+sh2.getName()+"!$B$3,\"' group by M\")),1)")
sh3.getRange(2,33,1,1).setValue("=QUERY($A:$O,IF("+sh2.getName()+"!$B$3=\"\",CONCATENATE(\"select avg(G)\"),CONCATENATE(\"select avg(G) where A contains'\","+sh2.getName()+"!$B$3,\"'\")),1)")
sh3.getRange(2,34,1,1).setValue("=QUERY($A:$O,IF("+sh2.getName()+"!$B$3=\"\",CONCATENATE(\"select sum(E)/sum(C) \"),CONCATENATE(\"select sum(E)/sum(C) where A contains'\","+sh2.getName()+"!$B$3,\"'\")),1)")
sh3.getRange(2,35,1,1).setValue("=QUERY($A:$O,IF("+sh2.getName()+"!$B$3=\"\",CONCATENATE(\"select sum(C)/sum(D) \"),CONCATENATE(\"select sum(C)/sum(D) where A contains'\","+sh2.getName()+"!$B$3,\"'\")),1)")
sh3.getRange(2,36,1,1).setValue("=QUERY($A:$O,IF("+sh2.getName()+"!$B$3=\"\",CONCATENATE(\"select G,Count(A),sum(E)/sum(C),sum(E)/sum(J) group by G\"),CONCATENATE(\"select G,count(A),sum(E)/sum(C),sum(E)/sum(J) where A contains'\","+sh2.getName()+"!$B$3,\"' group by G\")),1)")

//Hide the RAW data sheet
sh3.hideSheet();

// Copy the exported sheet data in the Main(Q_Score) Sheet   
var q=[["Avg. QS ---->","="+sh3.getName()+"!AG3"],["Avg. CPC ---->","="+sh3.getName()+"!AH3"],["Avg.CTR---->","="+sh3.getName()+"!AI3"]]  
var p=[["Q-Score Params","Above average","Average","Below average"],
       ["Ad relevance","=iferror(VLOOKUP(B$8,"+sh3.getName()+"!$AA:$AB,2,false),0)","=iferror(VLOOKUP(C$8,"+sh3.getName()+"!$AA:$AB,2,false),0)","=iferror(VLOOKUP(D$8,"+sh3.getName()+"!$AA:$AB,2,false),0)"],
      ["Exp. CTR","=iferror(VLOOKUP(B$8,"+sh3.getName()+"!$AC:$AD,2,false),0)","=iferror(VLOOKUP(C$8,"+sh3.getName()+"!$AC:$AD,2,false),0)","=iferror(VLOOKUP(D$8,"+sh3.getName()+"!$AC:$AD,2,false),0)"],
       ["Landing Page","=iferror(VLOOKUP(B$8,"+sh3.getName()+"!$AE:$AF,2,false),0)","=iferror(VLOOKUP(C$8,"+sh3.getName()+"!$AE:$AF,2,false),0)","=iferror(VLOOKUP(D$8,"+sh3.getName()+"!$AE:$AF,2,false),0)"]
     ]
var r=[["Quality Score","CPA","CPC"],
 ["1","=if(iferror(VLOOKUP($A23,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A23,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A23,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A23,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
 ["2","=if(iferror(VLOOKUP($A24,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A24,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A24,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A24,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
 ["3","=if(iferror(VLOOKUP($A25,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A25,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A25,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A25,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
  ["4","=if(iferror(VLOOKUP($A26,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A26,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A26,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A26,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
  ["5","=if(iferror(VLOOKUP($A27,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A27,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A27,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A27,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
  ["6","=if(iferror(VLOOKUP($A28,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A28,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A28,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A28,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
  ["7","=if(iferror(VLOOKUP($A29,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A29,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A29,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A29,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
  ["8","=if(iferror(VLOOKUP($A30,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A30,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A30,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A30,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
  ["9","=if(iferror(VLOOKUP($A31,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A31,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A31,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A31,"+sh3.getName()+"!$AJ:$AM,3,false),0))"],
  ["10","=if(iferror(VLOOKUP($A32,"+sh3.getName()+"!$AJ:$AM,4,false),0)=\"\",0,iferror(VLOOKUP($A32,"+sh3.getName()+"!$AJ:$AM,4,false),0))",
  "=if(iferror(VLOOKUP($A32,"+sh3.getName()+"!$AJ:$AM,3,false),0)=\"\",0,iferror(VLOOKUP($A32,"+sh3.getName()+"!$AJ:$AM,3,false),0))"]
 ]

  //Set all vlookup values
  sh2.getRange(4,1,3,2).setValues(q)
  sh2.getRange(8,1,4,4).setValues(p)
  sh2.getRange(21,1,1,1).setValue("Quality Score vs CPA vs CPC")
  sh2.getRange(22,1,11,3).setValues(r)

//Set the Format 
sh2.getRange("A21").setFontWeight("bold").setFontSize(12);
sh2.getRange("B4:B5").setNumberFormat("0.0");
sh2.getRange("B6").setNumberFormat("0.0%");
sh2.getRange("C3").setValue("=concatenate(SUM(B9:D9),\" \",\"Keywords\")")
sh2.getRange("B23:C32").setNumberFormat("0.0");

// Display the Date Range of the Data and Spreadhseet URL
Logger.log("Reports Created for Date Range "+start+" "+end)
Logger.log("URL:"+sh2.getParent().getUrl())
}


//Get Spreadhsheet 
function getSpreadsheetURL(name)
{
  var files = DriveApp.searchFiles('title contains "'+name+'"');
  if(files.hasNext()){
     var file = files.next();
 return file.getUrl();
  } else 
 {
var sh_new=SpreadsheetApp.create(name)
return sh_new.getUrl()
 }
}

//Get Sheet
function getsheet(sht,name){
  var sh2 =sht.getSheetByName(name);
  if(sh2)
   {
    return sht.getSheetByName(name)}
    else 
    { var sh2=sht.insertSheet(name)
           return sh2 
    }  
}

// Draw the Stacked Chart
function Charts_(sh2)
{
  var qchart=sh2.newChart();
  qchart.addRange(sh2.getRange("A8:D11")).setChartType(Charts.ChartType.COLUMN).asColumnChart().setStacked().setTitle("Quality Score").setPosition(2,5,50,0).setOption('width', 800);
  sh2.insertChart(qchart.build())

var qchart2=sh2.newChart();
qchart2.addRange(sh2.getRange("A22:B32")).setChartType(Charts.ChartType.COLUMN).asColumnChart().setTitle("Quality Score vs CPA").setPosition(21,5,50,0).setOption('width', 800);
  sh2.insertChart(qchart2.build())

}