from common_variables import *
from Report_Common_funcs import *
import os
import xml.etree.ElementTree as ET
import re
from Report_Common_funcs import *
global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
global j,i,Total_Test_CaseIn_Each_Scenario,l,TC_Not_Executed
global username,SSO,Serial,OS,CA
global RemaningScenario
j=0 #TestCase Lists Counter
i=0 #Scenario LIsts Counter
##------------------------------------------------------------------------------
##Purpose: Gets Unique ID to be used anywhere based on date-time.
##Author: Daipayan            Date:5/3/2019 
##-------------------------------------------------------------------------------
                
                
                
def Get_unique_id():
  CurrentDate = aqDateTime.Today();
  CurrentTime = aqDateTime.Time();
                
#  // Return the parts of the current date&time
  Year = aqDateTime.GetYear(CurrentDate);
  Month = aqDateTime.GetMonth(CurrentDate);
  Day = aqDateTime.GetDay(CurrentDate);
  Hours = aqDateTime.GetHours(CurrentTime);
  Minutes = aqDateTime.GetMinutes(CurrentTime);
                  
                
  if(Month<10):
    Month = "0" + Month
                    
  if(Day<10):
    Day = "0" + Day
                
  if(Hours<10):
    Hours = "0" + Hours
                    
  if(Minutes<10):
    Minutes = "0" + Minutes
                
                
  URID = Month+""+Day+""+Hours+""+Minutes
                  
  return URID
                
                
##------------------------------------------------------------------------------
##Purpose: Gets the current tested Req name.
##Author: Daipayan            Date:8/3/2019 
##-------------------------------------------------------------------------------
                
def Get_current_tested_req_name():
                
  try:
    TestItems = Project.TestItems;
    Log.Message("The " + TestItems.Current.Name + " test item is currently running.",TestItems.Current.Description);
    return TestItems.Current.Name;
                  
  except Exception as e:
    Log.Message(str(e))
                
##------------------------------------------------------------------------------
##Purpose: Gets the current tested Item description
##Author: Daipayan            Date:8/3/2019 
##-------------------------------------------------------------------------------
                
def Get_current_testitem_desc():
                
  try:
    TestItems = Project.TestItems;
    Log.Message("The " + TestItems.Current.Name + " test item is currently running.",TestItems.Current.Description);
                      
    return   TestItems.Current.Description;
  except Exception as e:
    Log.Message(str(e))
##------------------------------------------------------------------------------
##Purpose: Gets the Machine and User details from the Text File
##Author: Daipayan            Date:8/3/2019 
##-------------------------------------------------------------------------------
                
def get_user_details(userDetail,machineDetail):
  from Report_Common_funcs import readFile
  Filestr=getFile(User_details)
  username = readFile(Filestr,userDetail,"Name")
  SSO= readFile(Filestr,userDetail,"SSOID")
  Serial = readFile(Filestr,machineDetail,"Serial Number")
  OS = readFile(Filestr,machineDetail,"OS")
  CA = readFile(Filestr,machineDetail,"CA")
  SV= readFile(Filestr,machineDetail,"SV")
  return username,SSO,Serial,OS,CA,SV
                
##------------------------------------------------------------------------------
##Purpose: Invokes and Reads the XML files, reads data and prints the Header in HTML file.
##Author: Daipayan            Date:11/3/2019 
##-------------------------------------------------------------------------------
                
def Header_HTMLReport():
#      ReqID=Project.Variables.ReqID
      global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
      global j,i,Total_Test_CaseIn_Each_Scenario,LengthofScenarios
      global username,SSO,Serial,OS,CA,SV, HTML_Report_path
#      Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list =xml(Project.Variables.ReqID)#ReqID
      HTML_Report_path = HTML_Report_filepath +Req_ID+"-"+aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%b_%Y_%H_%M_%S")+'.html'
#      HTML_Report_path = Req_ID+"-"+aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%b_%Y_%H_%M_%S")+'.html'
      j=0
      username,SSO,Serial,OS,CA,SV = get_user_details("Daipayan","Machine1")
                      
#      LengthofScenarios=len(scenarios)
#      if len(scenarios)==len(Testcases_list):
                        
                      
      if(Utilities.FileExists(HTML_Report_path) == False):
          aqFile.Create(HTML_Report_path)
      Html_Styling()    
      Report_Header()
      Body_Req_details()
      Body_Environment_details()
                
##------------------------------------------------------------------------------
##Purpose: Prints the SCenario details in HTML file, takes scenario ID as parameter.
##Author: Daipayan            Date:3/4/2019 
##-------------------------------------------------------------------------------
                        
                
                
def print_Scenario():    
  global i,l,x
  if scenarios[i].get('Automate')=='Yes':
      Body_Scenario_details()
      Log.Message("Scenario {} is printed to HTML file :" .format(scenarios[i].get('ID')))
      return  
#def print_Scenario(ScenarioID):  
#  global i,j,LengthofScenarios,Total_Test_CaseIn_Each_Scenario
#
#  for i in range(len(scenarios)):
#    if (scenarios[i].get('ID'))== ScenarioID :
#      Body_Scenario_details() 
#      Total_Test_CaseIn_Each_Scenario=Testcases_list[i]                   
#      break
#  i=0
#  for i in range(len(scenarios)):
#    if (scenarios[i].get('ID'))== ScenarioID :
#      while i<len(Testcases_list):
#        if (Testcases_list[i]!=Total_Test_CaseIn_Each_Scenario):
#          j+=Testcases_list[i]
#          i+=1
#        else:
#          break
#        Log.Message("The value of j is :" +str(j))
                                           
                
                
##------------------------------------------------------------------------------
##Purpose: Prints the Test case details in HTML file.
##Author: Daipayan            Date:11/3/2019 
##-------------------------------------------------------------------------------
def Print_testCase(Flag):
                  
  global TestCasesInThisScenario,j,TCExecutedInThisScen, i,j
#  for j in range(TestCasesInThisScenario):
  Body_TC_details(Flag)
                    
  return
                                     
                
#def Print_testCase(Flag):
#  global Total_Test_CaseIn_Each_Scenario,j
#  
#  for tc in range(Total_Test_CaseIn_Each_Scenario):
#    Body_TC_details(Flag)
#    
#    j=j+1
#    return
                  
##------------------------------------------------------------------------------
##Purpose: Prints the CSS Part of the HTML file in the Header section.
##Author: Daipayan            Date:12/3/2019 
##-------------------------------------------------------------------------------
                                                                                          
def Html_Styling():
          global HTML_Report_path
          oFile = aqFile.OpenTextFile(HTML_Report_path, aqFile.faWrite, aqFile.ctUTF8, True)
          oFile.WriteLine("<!DOCTYPE html> ")
          oFile.WriteLine("<html>")
          oFile.WriteLine("<head>")
          oFile.WriteLine("<style>")
          oFile.WriteLine(".header img {float: left;width: 70px;height: 70px;background: #555;}")
          oFile.WriteLine(".header h1 {text-align:center;top: 18px;left: 10px;}")
          oFile.WriteLine("table {")
          oFile.WriteLine("width:100%;")
          oFile.WriteLine("}")
          oFile.WriteLine("table, th, td {")
          oFile.WriteLine("border: 1px solid black;")
          oFile.WriteLine("border-collapse: collapse;")
          oFile.WriteLine("}")
          oFile.WriteLine("th, td {")
          oFile.WriteLine("padding: 5px;")
          oFile.WriteLine("text-align: left;")
          oFile.WriteLine("}")
          oFile.WriteLine("table.names tr:nth-child(even) { ")
          oFile.WriteLine("background-color: #C0C0C0; ")
          oFile.WriteLine("} ")
          oFile.WriteLine("table.names tr:nth-child(odd) { ")
          oFile.WriteLine("background-color:#C0C0C0; ")
          oFile.WriteLine("} ")
          oFile.WriteLine("table.names th { ")
          oFile.WriteLine("background-color: #FFFFFF; }")
          oFile.WriteLine(".Mytbl {width:100%;background-color: #87CEEB;} ")
          oFile.WriteLine(".Mytblfooter {width:100%;background-color: #ffffff;} ")
          oFile.WriteLine(".Mytblconfig {width:100%;background-color: #B0B7F0;}")
          oFile.WriteLine("img[alt='screenshot']{width:95%;height:95%}")
                          
          oFile.WriteLine(".dot{height:150px;width: 150px;border-radius: 50%; display:inline-block; background-color:#32CD32;}")
          oFile.WriteLine("p {display: block;padding: 30px;}")
          oFile.WriteLine(".ReqStatus td{padding-left: 20px;align:center;border: 0px solid white;font-size:30px;}")
          oFile.WriteLine(".ReqStatus{border: 0px solid white;")
                
                          
          oFile.WriteLine("</style> ")
          oFile.WriteLine("</head> ")
          oFile.WriteLine("<body> ")
          oFile.Close()
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the REport Header.
##Author: Daipayan            Date:12/3/2019 
##-------------------------------------------------------------------------------
                          
                                                                                     
def Report_Header():
    global username,SSO,Serial,OS,CA,SV, HTML_Report_path
    aqFile.WriteToTextFile(HTML_Report_path, "<div class='header'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, '<img src="C:\\Work\\CSCS\\CSCS_Automation\\CSCS_Automation\\InputData\\Logo\\GE_Logo.jpg" alt="logo" />',aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<h1>" +Req_ID+  ": Execution Report</h1>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</div><br/>",aqFile.ctUTF8)
                    
##------------------------------------------------------------------------------
##Purpose: HTML code of the Requirement part in the HTML Body 
##Author: Daipayan            Date:13/3/2019 
##-------------------------------------------------------------------------------
                    
def Body_Req_details():
                    
    global username,SSO,Serial,OS,CA,SV, HTML_Report_path
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='Mytbl' >",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement Name</td><td width='55%'>" +Req_Desc+ "</td><td style='text-align:center'>Status</td></tr> ",aqFile.ctUTF8)
    Print_Req_Execution_status(Project.Variables.req_status)                                                                   
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>SSO ID</td><td>" +SSO+ "</td>" + "</tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Executed By</td><td>" +username+ "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Date & Time Executed</td><td>" + str(aqDateTime.Now()) + "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the UUT Details part in the HTML Body 
##Author: Daipayan            Date:13/3/2019 
##-------------------------------------------------------------------------------
                   
def Body_Environment_details():
    global username,SSO,Serial,OS,CA,SV, HTML_Report_path                                                                      
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='Mytbl'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='25%' colspan='4' style='text-align:center'>UUT Details</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Serial Number</td><td>CA</td><td> OS:</td><td> SV:</td></tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td height='30'>"+Serial+"</td><td>"+CA+"</td><td>"+OS+" </td><td>"+SV+" </td></tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</table>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the Scenario Details part in the HTML Body 
##Author: Daipayan            Date:13/3/2019 
##-------------------------------------------------------------------------------
                                                                                        
def Body_Scenario_details():
    global i, HTML_Report_path
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='names' style='font-weight: bold'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='25%'>Scenario ID</td><td>" +scen_ID_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Scenario Name</td><td>" +scen_Desc_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    Body_Scenario_Precondition()
    aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the Scenario Precondtion part in the HTML Body. 
##Prints only if available in XML. 
##Author: Daipayan            Date:15/3/2019 
##-------------------------------------------------------------------------------
                    
def Body_Scenario_Precondition():
    global i, HTML_Report_path
    if SPreconditon_list[i]!= None:
      aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Scenario Precondtion</td><td>" +SPreconditon_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    else:
      pass
##------------------------------------------------------------------------------
##Purpose: HTML code of the TC details in the HTML Body. 
##Author: Daipayan            Date:15/3/2019 
##-------------------------------------------------------------------------------
                                                                                    
def Body_TC_details(Flag):
    global j,i, HTML_Report_path
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='names' style='font-weight: normal'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='8%'>TestCase ID</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>Action</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>Expected result</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>Actual result</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='10%'>Status</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='8%'>"+TC_ID_list[j]+"</td>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='30%'>"+TC_Action_list[j]+"</td>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='30%'>"+TC_Expected_list[j]+"</td>  ",aqFile.ctUTF8)
    print_Actual(Flag)
    aqFile.WriteToTextFile(HTML_Report_path, "</tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
#    take_Screenshot()
    print_screenshot_to_HTML()
                
##------------------------------------------------------------------------------
##Purpose: HTML code for printing the Pass/Fail staus of Test Case in the HTML Body. 
##Prints only if available in XML. 
##Author: Daipayan            Date:19/3/2019 
##-------------------------------------------------------------------------------
                    
                
def print_Actual(Flag):
  global j,i, HTML_Report_path
  if (Flag=='Pass'):
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='30%'>"+TC_ActualPass_list[j]+"</td>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='10%' style = 'background-color: #32CD32' data-test-status='tcstatus' >Pass</td>  ",aqFile.ctUTF8)
  elif Flag=='Fail':
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='30%'>"+TC_ActualFail_list[j]+"</td>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='10%' style = 'background-color: #FF0000' data-test-status='tcstatus'>Fail</td>  ",aqFile.ctUTF8)
  else:
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='30%'>""</td>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<td width='10%'>Not Executed</td>  ",aqFile.ctUTF8)
                
##------------------------------------------------------------------------------
##Purpose: HTML code to complete the HTML code. 
##Also it prints if any Test cases part of any scenario is not executed 
##Prints only if available in XML. 
##Author: Daipayan            Date:28/3/2019 
##-------------------------------------------------------------------------------
                    
def on_stop():
   global Total_Test_CaseIn_Each_Scenario,j,TC_Not_Executed
   TC_Not_Executed=Total_Test_CaseIn_Each_Scenario - Project.Variables.TC_executed
   if TC_Not_Executed!=0:
      aqFile.WriteToTextFile(HTML_Report_path, "<p style='font-size: 35px;color:#ff0000'>TC Remaining to be executed within this Scenario: "+(str(TC_Not_Executed))+"</p>",aqFile.ctUTF8)
                       
   aqFile.WriteToTextFile(HTML_Report_path, "</body>",aqFile.ctUTF8)
   aqFile.WriteToTextFile(HTML_Report_path, "</html>",aqFile.ctUTF8)
   ConvertHTMLtoPDF()
                  
##------------------------------------------------------------------------------
##Purpose: Takes screenshot and places it in appropiate directory. 
##Also it has The HTML code to mention the src of the screenshot.
##Prints only if available in XML. 
##Author: Daipayan            Date:19/3/2019 
##-------------------------------------------------------------------------------
                  
def take_Screenshot(ScenarioID):
  try:
                    
    global j,i,HTML_Report_path,img_name,screenshot_list,screenshot_list
    global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
    req_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y")
    screenshot_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y_%I_%M_%S")
    Foldername= Req_ID+'_'+ req_dateInfo
    img_dir = Project.Path+'Screenshots\\'+Foldername
#    j=0
#    screenshot_filename = Req_ID+TC_ID_list[j]+screenshot_dateInfo
                    
    #adding file name
                    
#    for j in range(len(TC_ID_list)):
#        Scenario_ID=TC_ID_list[j].split('_')
#        if ScenarioID in Scenario_ID[0]:
#            screenshot_filename = Req_ID+TC_ID_list[j]+screenshot_dateInfo
#            if j==0:
#        #    if i==0:
#              if(Utilities.DirectoryExists(img_dir)==True):
#                Log.Message("Screenshot_Req Folder with same name already exists :" + Foldername)
#              else:
#                aqFileSystem.CreateFolder(img_dir)
#                Log.Message("Screenshot_Req Folder created" +Foldername)
    screenshot_name = set_screenshot_filename(ScenarioID)
    if(Utilities.FileExists(img_dir+'\\'+screenshot_name+'.jpg') == False):
                  pic = Sys.Desktop
                  img_name= img_dir+'\\'+screenshot_name+'.jpg'
                  pic.Picture().SaveToFile(img_name)
                  Log.Message("screenshot File created :" +screenshot_name)
                  screenshot_list.append(screenshot_name)
#                  j=j+1
    else:
                 Log.Message("Screenshot not created")
  except Exception as e:
    Log.Message(str(e))
                    
def set_screenshot_filename(ScenarioID):
    global j,i,HTML_Report_path,img_name,counter,img_dir
    global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
    req_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y")
    screenshot_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y_%I_%M_%S")
    Foldername= Req_ID+'_'+ req_dateInfo
    img_dir = Project.Path+'Screenshots\\'+Foldername
                    
#    screenshot_filename = Req_ID+TC_ID_list[j]+screenshot_dateInfo
                    
    #adding file name
                    
    for itr in range(len(TC_ID_list)):
        Scenario_ID=TC_ID_list[itr].split('_')
        if ScenarioID in Scenario_ID[0]:
            screenshot_filename = Req_ID+'_'+TC_ID_list[counter]+'-'+screenshot_dateInfo
            if counter==0:
        #    if i==0:
              if(Utilities.DirectoryExists(img_dir)==True):
                Log.Message("Screenshot_Req Folder with same name already exists :" + Foldername)
              else:
                aqFileSystem.CreateFolder(img_dir)
                Log.Message("Screenshot_Req Folder created" +Foldername)
            counter=counter+1
            return screenshot_filename
                
                
                    
def print_screenshot_to_HTML():
  global j
                  
  screenshot_name = screenshot_list[j]
  img_name= img_dir+'\\'+screenshot_name+'.jpg'
  aqFile.WriteToTextFile(HTML_Report_path, '<img src="%s" alt="screenshot" />'%(img_name),aqFile.ctUTF8)
  Log.Message(screenshot_name+ " screenshot written to HTML" )
                
##------------------------------------------------------------------------------
##Purpose: Prints the Overall Requirement status.
##Author: Daipayan            Date:27/3/2019 
##-------------------------------------------------------------------------------
                    
def Print_Req_Execution_status(req_status):
  global RemaningScenario
  Project.Variables.req_status=req_status
  #Updated by  (reversed the if and elif conditions)
  if req_status=='Fail' : #or Project.Variables.TC_executed !=0
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement PreCondition</td><td width='55%'>" +Precondition+ "</td><td  id='req' rowspan ='5' style='text-align:center;background-color:#ff0000'>Fail</td></tr> ",aqFile.ctUTF8)
          
  elif req_status=='Pass' : #and Project.Variables.TC_executed ==0 and RemaningScenario==0
      aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement PreCondition</td><td width='55%'>" +Precondition+ "</td><td  id='req' rowspan ='5' style='text-align:center; background-color:#008000'>Pass</td></tr> ",aqFile.ctUTF8)
                
#  if req_status=='Fail' or Project.Variables.TC_executed !=0 or RemaningScenario!=0:
#    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement PreCondition</td><td width='55%'>" +Precondition+ "</td><td  id='req' rowspan ='5' style='text-align:center;background-color:#ff0000'>Fail</td></tr> ",aqFile.ctUTF8)
#
#  elif req_status=='Pass':
#    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement PreCondition</td><td width='55%'>" +Precondition+ "</td><td  id='req' rowspan ='5' style='text-align:center; background-color:#008000'>Pass</td></tr> ",aqFile.ctUTF8)
                
##------------------------------------------------------------------------------
##Purpose: Convert the HTML Report to PDF file using a dLL
##Author: Daipayan            Date:27/3/2019 
##-------------------------------------------------------------------------------
                
def ConvertHTMLtoPDF():
  global j,i,HTML_Report_path,Req_ID
  pdf_timestamp= aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d-%b-%Y-%H_%M_%S")
  pdf_filename=Req_ID+'_'+pdf_timestamp
  pdf_fullname= Project.Path +'Reports\\Pdf\\' +pdf_filename+'.pdf'
  pdf_convert = dotNET.HtmlToPdfUtility.HtmlToPdfConverter.CreatePdfFromHtmlFile(HTML_Report_path,pdf_fullname)
  Log.Message("Status of HTML to PDF conversion: " +str(pdf_convert))
                
                
              
from collections import OrderedDict
Scen_dict=OrderedDict()
def HTML_Report_print():
  global Scen_dict,TestCasesInThisScenario,TCExecutedInThisScen, i,j, RemaningScenario
  j=0 #Counter for TestCase LIsts
  FailedReq=False
#  Project.Variables.TC_executed=len(TCStatus)
                 
  #The Below code gets the Overall Req Status and stores it in global variable
  for key,val in Scen_dict.items():
                     
    if False in val:
      FailedReq=True
      break
                    
  if FailedReq==True:
      Project.Variables.req_status='Fail'
  else:
      Project.Variables.req_status='Pass'
#  Header_HTMLReport()
  ScenariosAvailableInThisReq = len(scenarios)
  ScenariosExecutedInThisReq = len(Scen_dict)
  if ScenariosAvailableInThisReq != ScenariosExecutedInThisReq :
       RemaningScenario= ScenariosAvailableInThisReq - ScenariosExecutedInThisReq
       if RemaningScenario !=0:
            Project.Variables.req_status='Fail'
        
  #The below code iterates through the XML and gets the TC count available for the mentioned Scenario
  for i in range(len(scen_ID_list)):
#    if scen_ID_list[iter] == '01':#Scen_ID
              TestCasesInThisScenario = Testcases_list[i]
                              
              #The below code iterates through the dict and gets the TC count executed for the mentioned Scenario
              for key,val in Scen_dict.items():
                                         
                          if key==scen_ID_list[i]:##Scen_ID
                                  TCExecutedInThisScen = len(val) 
                                                                                   
                                  #Looks for any difference in TC count in XML vs actually executed
                                  if TCExecutedInThisScen!= TestCasesInThisScenario:
                                        RemainingTC = TestCasesInThisScenario - TCExecutedInThisScen
                                        if RemainingTC!=0:
                                              Project.Variables.req_status='Fail'
          
          
  Header_HTMLReport()
          
          
          
  for i in range(len(scen_ID_list)):
#    if scen_ID_list[iter] == '01':#Scen_ID
              TestCasesInThisScenario = Testcases_list[i]
                              
              #The below code iterates through the dict and gets the TC count executed for the mentioned Scenario
              for key,val in Scen_dict.items():
                                         
                          if key==scen_ID_list[i]:##Scen_ID
                                  TCExecutedInThisScen = len(val) 
                                  print_Scenario()
                                  for itr in range(len(val)):
                                        if val[itr]==True:
                                              Flag='Pass'
                                        else:
                                              Flag='Fail'
                                                         
                                        Print_testCase(Flag)
                                        j+=1
                                                        
                                  #Looks for any difference in TC count in XML vs actually executed
                                  if TCExecutedInThisScen!= TestCasesInThisScenario:
                                        RemainingTC = TestCasesInThisScenario - TCExecutedInThisScen
                                        if RemainingTC!=0:
                                              aqFile.WriteToTextFile(HTML_Report_path, "<p style='font-size: 35px;color:#ff0000'>TC Remaining to be executed within this Scenario: "+(str(RemainingTC))+"</p>",aqFile.ctUTF8)
  if ScenariosAvailableInThisReq != ScenariosExecutedInThisReq :
     RemaningScenario= ScenariosAvailableInThisReq - ScenariosExecutedInThisReq
     if RemaningScenario !=0:
          aqFile.WriteToTextFile(HTML_Report_path, "<p style='font-size: 35px;color:#ff0000'>Scenario Remaining to be executed within this Requirement: "+(str(RemaningScenario))+"</p>",aqFile.ctUTF8)   
                                          
#                          break         
  aqFile.WriteToTextFile(HTML_Report_path, "</body>",aqFile.ctUTF8)
  aqFile.WriteToTextFile(HTML_Report_path, "</html>",aqFile.ctUTF8)
  ConvertHTMLtoPDF()
  onStop_TestedItems()
                
                  
def On_start_Req(ReqID):
  global Scen_dict,TestCasesInThisScenario,counter,screenshot_list
  global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
  Project.Variables.ReqID=ReqID
  counter=0
  screenshot_list=[]
  Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list =xml(Project.Variables.ReqID)#ReqID
                  
def TC_validation_status(ScenarioID,TC_Status):
  global Scen_dict,TestCasesInThisScenario
  Scen_dict.setdefault(ScenarioID, []).append(TC_Status)
#  take_Screenshot()
  take_Screenshot(ScenarioID)
                  
def onStop_TestedItems():
  global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
  global Scen_dict,TestCasesInThisScenario,counter,screenshot_list,j,i
  scenarios_list=[]
  scen_ID_list=[]
  scen_Desc_list=[]
  Scen_dict={}
  SPreconditon_list =[]
  Testcases_list=[]
  TC_ID_list=[]
  TC_Action_list=[]
  TC_Expected_list =[]
  TC_ActualPass_list=[]
  TC_ActualFail_list =[]
  TestCasesInThisScenario=[]               
  screenshot_list=[]               
  j=0
  i=0
  counter=0   
  Req_ID=''           
  Req_Desc=''       
  Precondition=''
                  
                
                    
                      
