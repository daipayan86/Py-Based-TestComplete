
from common_variables import *
from Report_Common_funcs import *
import os
import xml.etree.ElementTree as ET
import re
from Report_Common_funcs import *
global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
global j,i,Total_Test_CaseIn_Each_Scenario,l,TC_Not_Executed
global username,SSO,Serial,OS,CA
global RemaningScenario
global Str_exception
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
    save_exception(e)
                
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
  try:
      from Report_Common_funcs import readFile
      Filestr=getFile(User_details)
      username = readFile(Filestr,userDetail,"Name")
      SSO= readFile(Filestr,userDetail,"SSOID")
      Serial = readFile(Filestr,machineDetail,"Serial Number")
      OS = readFile(Filestr,machineDetail,"OS")
      CA = readFile(Filestr,machineDetail,"CA")
      SV= readFile(Filestr,machineDetail,"SV")
      return username,SSO,Serial,OS,CA,SV
  except Exception as e:
      save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: Invokes and Reads the XML files, reads data and prints the Header in HTML file.
##Author: Daipayan            Date:11/3/2019 
##-------------------------------------------------------------------------------
                
def Header_HTMLReport():
  try:
      global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
      global j,i,Total_Test_CaseIn_Each_Scenario,LengthofScenarios
      global username,SSO,Serial,OS,CA,SV, HTML_Report_path
      HTML_Report_path = HTML_Report_filepath +Req_ID+"-"+aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%b_%Y_%H_%M_%S")+'.html'
      j=0
      username,SSO,Serial,OS,CA,SV = get_user_details("Username","Machine_Details")
                   
                      
      if(Utilities.FileExists(HTML_Report_path) == False):
          aqFile.Create(HTML_Report_path)
      Html_Styling()    
      Report_Header()
      Body_Req_details()
      Body_Environment_details()
  except Exception as e:
    save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: Prints the SCenario details in HTML file, takes scenario ID as parameter.
##Author: Daipayan            Date:3/4/2019 
##-------------------------------------------------------------------------------
                        
                
                
def print_Scenario():
  try:    
      global i,l,x
      if scenarios[i].get('Automate')=='Yes':
          Body_Scenario_details()
          Log.Message("Scenario {} is printed to HTML file :" .format(scenarios[i].get('ID')))
          return  
  except Exception as e:
      save_exception(e)

                
                
##------------------------------------------------------------------------------
##Purpose: Prints the Test case details in HTML file.
##Author: Daipayan            Date:11/3/2019 
##-------------------------------------------------------------------------------
def Print_testCase(Flag):
  try:             
      global TestCasesInThisScenario,j,TCExecutedInThisScen, i,j
    #  for j in range(TestCasesInThisScenario):
      Body_TC_details(Flag)            
      return
  except Exception as e:
      save_exception(e)
                                     
                
                  
##------------------------------------------------------------------------------
##Purpose: Prints the CSS Part of the HTML file in the Header section.
##Author: Daipayan            Date:12/3/2019 
##-------------------------------------------------------------------------------
                                                                                          
def Html_Styling(*args):
  try:
          global HTML_Report_path,HTML_Summary_Report_path
          if len(args)==0:
              Type_report_path = HTML_Report_path
          else:
              Type_report_path = HTML_Summary_Report_path
          oFile = aqFile.OpenTextFile(Type_report_path, aqFile.faWrite, aqFile.ctUTF8, True)
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
          oFile.WriteLine(".TCNotExecuted {width:30%;background-color: #fd9494;align:center} ")
          oFile.WriteLine(".Mytblfooter {width:100%;background-color: #ffffff;} ")
          oFile.WriteLine(".Mytblconfig {width:100%;background-color: #ffffff;}")
          oFile.WriteLine("img[alt='screenshot']{width:95%;height:95%}")              
          oFile.WriteLine(".dot{height:150px;width: 150px;border-radius: 50%; display:inline-block; background-color:#32CD32;}")
          oFile.WriteLine("p {display: block;padding: 30px;}")
          oFile.WriteLine(".ReqStatus td{padding-left: 20px;align:center;border: 0px solid white;font-size:30px;}")
          oFile.WriteLine(".ReqStatus{border: 0px solid white;")
          oFile.WriteLine("</style> ")
          oFile.WriteLine("</head> ")
          oFile.WriteLine("<body> ")
          oFile.Close()
  except Exception as e:
          save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the REport Header.
##Author: Daipayan            Date:12/3/2019 
##-------------------------------------------------------------------------------
                          
                                                                                     
def Report_Header(*args):
#  try:
#    global username,SSO,Serial,OS,CA,SV, HTML_Report_path
#    aqFile.WriteToTextFile(HTML_Report_path, "<div class='header'>",aqFile.ctUTF8)
#    aqFile.WriteToTextFile(HTML_Report_path, '<img src="C:\\Work\\CSCS\\CSCS_Automation\\CSCS_Automation\\InputData\\Logo\\GE_Logo.jpg" alt="logo" />',aqFile.ctUTF8)
#    if args[0].lower()=='Summary'.lower():
#      aqFile.WriteToTextFile(HTML_Report_path, "<h1> Execution Summary Report</h1>",aqFile.ctUTF8)
#    else:
#      aqFile.WriteToTextFile(HTML_Report_path, "<h1>" +Req_ID+  ": Execution Report</h1>",aqFile.ctUTF8)
#    aqFile.WriteToTextFile(HTML_Report_path, "</div><br/>",aqFile.ctUTF8)
    
  try:
    global username,SSO,Serial,OS,CA,SV, HTML_Report_path,HTML_Summary_Report_path
    if len(args)==0:
      Type_report_path = HTML_Report_path
    else:
      Type_report_path = HTML_Summary_Report_path
        
    aqFile.WriteToTextFile(Type_report_path, "<div class='header'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, '<img src="C:\\Work\\CSCS\\CSCS_Automation\\CSCS_Automation\\InputData\\Logo\\GE_Logo.jpg" alt="logo" />',aqFile.ctUTF8)
    if len(args)==0:
          aqFile.WriteToTextFile(Type_report_path, "<h1>" +Req_ID+  ": Execution Report</h1>",aqFile.ctUTF8)
    elif args[0].lower()=='Summary'.lower():
          aqFile.WriteToTextFile(Type_report_path, "<h1> Execution Summary Report</h1>",aqFile.ctUTF8)
    
    aqFile.WriteToTextFile(Type_report_path, "</div><br/>",aqFile.ctUTF8)
  except Exception as e:
    save_exception(e)
                    
##------------------------------------------------------------------------------
##Purpose: HTML code of the Requirement part in the HTML Body 
##Author: Daipayan            Date:13/3/2019 
##-------------------------------------------------------------------------------
                    
def Body_Req_details(*args):
  try:                    
    global username,SSO,Serial,OS,CA,SV, HTML_Report_path,HTML_Summary_Report_path
    if len(args)==0:
            Type_report_path = HTML_Report_path
    else:
            Type_report_path = HTML_Summary_Report_path
    aqFile.WriteToTextFile(Type_report_path, "<table class='Mytbl' >",aqFile.ctUTF8)
    if len(args)==0:
        aqFile.WriteToTextFile(Type_report_path, "<tr><td width='20%'>Requirement Name</td><td width='55%'>" +Req_Desc+ "</td><td style='text-align:center'>Status</td></tr> ",aqFile.ctUTF8)
        Print_Req_Execution_status(Project.Variables.req_status)
    elif args[0].lower()=='Summary'.lower(): 
        aqFile.WriteToTextFile(Type_report_path, "<tr><td width='25%' colspan='2' style='text-align:center'><b>User Details</b></td></tr> ",aqFile.ctUTF8)                                                                 
    aqFile.WriteToTextFile(Type_report_path, "<tr><td>SSO ID</td><td>" +SSO+ "</td>" + "</tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<tr><td>Executed By</td><td>" +username+ "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<tr><td>Date & Time Executed</td><td>" + str(aqDateTime.Now()) + "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<br/> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<br/> ",aqFile.ctUTF8)
  except Exception as e:
    save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the UUT Details part in the HTML Body 
##Author: Daipayan            Date:13/3/2019 
##-------------------------------------------------------------------------------
                   
def Body_Environment_details(*args):
  try:
    global username,SSO,Serial,OS,CA,SV, HTML_Report_path,HTML_Summary_Report_path    
    if len(args)==0:
            Type_report_path = HTML_Report_path
    else:
            Type_report_path = HTML_Summary_Report_path                                                                  
    aqFile.WriteToTextFile(Type_report_path, "<table class='Mytbl'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<tr><td width='25%' colspan='4' style='text-align:center'>UUT Details</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<tr><td>Serial Number</td><td>CA</td><td> OS:</td><td> SV:</td></tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<tr><td height='30'>"+Serial+"</td><td>"+CA+"</td><td>"+OS+" </td><td>"+SV+" </td></tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "</table>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<br/> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(Type_report_path, "<br/> ",aqFile.ctUTF8)
  except Exception as e:
    save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the Scenario Details part in the HTML Body 
##Author: Daipayan            Date:13/3/2019 
##-------------------------------------------------------------------------------
                                                                                        
def Body_Scenario_details():
  try:
    global i, HTML_Report_path
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='names' style='font-weight: bold'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='25%'>Scenario ID</td><td>" +scen_ID_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Scenario Name</td><td>" +scen_Desc_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    Body_Scenario_Precondition()
    aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
  except Exception as e:
    save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: HTML code of the Scenario Precondtion part in the HTML Body. 
##Prints only if available in XML. 
##Author: Daipayan            Date:15/3/2019 
##-------------------------------------------------------------------------------
                    
def Body_Scenario_Precondition():
  try:
    global i, HTML_Report_path
    if SPreconditon_list[i]!= None:
      aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Scenario Precondtion</td><td>" +SPreconditon_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    else:
      pass
  except Exception as e:
    save_exception(e)
##------------------------------------------------------------------------------
##Purpose: HTML code of the TC details in the HTML Body. 
##Author: Daipayan            Date:15/3/2019 
##-------------------------------------------------------------------------------
                                                                                    
def Body_TC_details(Flag):
  try:
        global j,i, HTML_Report_path,Scen_dict

        aqFile.WriteToTextFile(HTML_Report_path, "<table class='names' style='font-weight: normal'>",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<tr> ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<th width='8%'>TestCase ID</th>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<th width='15%'>TestCase Description</th>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<th width='20%'>Action</th>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<th width='15%'>Expected result</th>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>Actual result</th>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<th width='10%'>Status</th>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "</tr>",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<tr> ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<td width='8%'>"+TC_ID_list[j]+"</td>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<td width='15%'>"+TC_Description_list[j]+"</td>  ",aqFile.ctUTF8)        
        aqFile.WriteToTextFile(HTML_Report_path, "<td width='20%'>"+TC_Action_list[j]+"</td>  ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<td width='15%'>"+TC_Expected_list[j]+"</td>  ",aqFile.ctUTF8)
        print_Actual(Flag)
        aqFile.WriteToTextFile(HTML_Report_path, "</tr>",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
        print_screenshot_to_HTML()
  except Exception as e:
        save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: HTML code for printing the Pass/Fail staus of Test Case in the HTML Body. 
##Prints only if available in XML. 
##Author: Daipayan            Date:19/3/2019 
##-------------------------------------------------------------------------------
                    
                
def print_Actual(Flag):
  try:
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
  except Exception as e:
    save_exception(e)
                
##------------------------------------------------------------------------------
##Purpose: HTML code to complete the HTML code. 
##Also it prints if any Test cases part of any scenario is not executed 
##Prints only if available in XML. 
##Author: Daipayan            Date:28/3/2019 
##-------------------------------------------------------------------------------
                    
def on_stop():
  try:
         global Total_Test_CaseIn_Each_Scenario,j,TC_Not_Executed
         TC_Not_Executed=Total_Test_CaseIn_Each_Scenario - Project.Variables.TC_executed
         if TC_Not_Executed!=0:
            aqFile.WriteToTextFile(HTML_Report_path, "<p style='font-size: 35px;color:#ff0000'>TC Remaining to be executed within this Scenario: "+(str(TC_Not_Executed))+"</p>",aqFile.ctUTF8)
                       
         aqFile.WriteToTextFile(HTML_Report_path, "</body>",aqFile.ctUTF8)
         aqFile.WriteToTextFile(HTML_Report_path, "</html>",aqFile.ctUTF8)
         
         ConvertHTMLtoPDF()
  except Exception as e:
         save_exception(e)
                  
##------------------------------------------------------------------------------
##Purpose: Takes screenshot and places it in appropiate directory. 
##Also it has The HTML code to mention the src of the screenshot.
##Prints only if available in XML. 
##Author: Daipayan            Date:19/3/2019 
##-------------------------------------------------------------------------------
                  
def take_Screenshot(TC_ID):
  try:                
    global j,i,HTML_Report_path,img_name,screenshot_list,screenshot_list
    global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
    req_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y")
    screenshot_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y_%I_%M_%S")
    Foldername= Req_ID+'_'+ req_dateInfo
    img_dir = Project.Path+'Screenshots\\'+Foldername

    if  set_screenshot_filename(TC_ID) == False:
      Log.Message('Recheck TC ID and XML details')
    else:
      screenshot_name = set_screenshot_filename(TC_ID)
      if(Utilities.FileExists(img_dir+'\\'+screenshot_name+'.jpg') == False):
                    pic = Sys.Desktop
                    img_name= img_dir+'\\'+screenshot_name+'.jpg'
                    pic.Picture().SaveToFile(img_name)
                    Log.Message("screenshot File created :" +screenshot_name)
                    screenshot_list.append(screenshot_name)

      else:
                    Log.Message("Screenshot not created")
  except Exception as e:
    save_exception(e)
    Log.Message(str(e))

                    

def set_screenshot_filename(TC_ID):
  try:
        global j,i,HTML_Report_path,img_name,counter,img_dir
        global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
        req_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y")
        screenshot_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y_%I_%M_%S")
        Foldername= Req_ID+'_'+ req_dateInfo
        img_dir = Project.Path+'Screenshots\\'+Foldername        
        if TC_ID in TC_ID_list:
                screenshot_filename = TC_ID +'-'+screenshot_dateInfo 
                if counter==0:
            #    if i==0:
                  if(Utilities.DirectoryExists(img_dir)==True):
                    Log.Message("Screenshot_Req Folder with same name already exists :" + Foldername)
                  else:
                    aqFileSystem.CreateFolder(img_dir)
                    Log.Message("Screenshot_Req Folder created" +Foldername)
                counter=counter+1
                return screenshot_filename
        else:
          return False
  except Exception as e:
      save_exception(e)
      Log.Message(e)



                
                
                    
def print_screenshot_to_HTML():
  try:
      global j,TC_ID_list
      for screenshot_name in screenshot_list:
          if TC_ID_list[j] in  screenshot_name:                
    #              screenshot_name = screenshot_list[j]
                  img_name= img_dir+'\\'+screenshot_name+'.jpg'
                  aqFile.WriteToTextFile(HTML_Report_path, '<img src="%s" alt="screenshot" />'%(img_name),aqFile.ctUTF8)
                  Log.Message(screenshot_name+ " screenshot written to HTML" )
  except Exception as e:
      save_exception(e)
      Log.Message(str(e))          
##------------------------------------------------------------------------------
##Purpose: Prints the Overall Requirement status.
##Author: Daipayan            Date:27/3/2019 
##-------------------------------------------------------------------------------
                    
def Print_Req_Execution_status(req_status):
  try:
        global RemaningScenario
        Project.Variables.req_status=req_status
        #Updated by  (reversed the if and elif conditions)
        if req_status=='Fail' : #or Project.Variables.TC_executed !=0
          aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement PreCondition</td><td width='55%'>" +Precondition+ "</td><td  id='req' rowspan ='5' style='text-align:center;background-color:#ff0000'>Fail</td></tr> ",aqFile.ctUTF8)
          
        elif req_status=='Pass' : #and Project.Variables.TC_executed ==0 and RemaningScenario==0
            aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement PreCondition</td><td width='55%'>" +Precondition+ "</td><td  id='req' rowspan ='5' style='text-align:center; background-color:#008000'>Pass</td></tr> ",aqFile.ctUTF8)
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))
                
##------------------------------------------------------------------------------
##Purpose: Convert the HTML Report to PDF file using a dLL
##Author: Daipayan            Date:27/3/2019 
##-------------------------------------------------------------------------------
                
def ConvertHTMLtoPDF():
  try:
    
   global j,i,HTML_Report_path,Req_ID
   pdf_timestamp= aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d-%b-%Y-%H_%M_%S")
   pdf_filename=Req_ID+'_'+pdf_timestamp
   pdf_fullname= Project.Path +'Reports\\Pdf\\' +pdf_filename+'.pdf'
   pdf_convert = dotNET.HtmlToPdfUtility.HtmlToPdfConverter.CreatePdfFromHtmlFile(HTML_Report_path,pdf_fullname)
   Log.Message("Status of HTML to PDF conversion: " +str(pdf_convert))
                # Commenting it for Test Purpose
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))  
                
              
from collections import OrderedDict
Scen_dict=OrderedDict()
def HTML_Report_print():
  try:
        global Scen_dict,TestCasesInThisScenario,TCExecutedInThisScen, i,j, RemaningScenario,HTML_Report_path
        global Str_exception
        j=0 #Counter for TestCase LIsts
        FailedReq=False
        if  len(Scen_dict) == 0:
            Header_HTMLReport()
            Print_Exception()
        else:
               
        #The Below code gets the Overall Req Status and stores it in global variable
                for key,val in Scen_dict.items():
                    for TC,Status in (Scen_dict.get(key)).items():
                      if (Scen_dict.get(key)).get(TC)==False:
                        FailedReq=True 
                        break                
                  
                if FailedReq==True:
                    Project.Variables.req_status='Fail'
                else:
                    Project.Variables.req_status='Pass'

                ScenariosAvailableInThisReq = len(scenarios)
                ScenariosExecutedInThisReq = len(Scen_dict)
                if ScenariosAvailableInThisReq != ScenariosExecutedInThisReq :
                     RemaningScenario= ScenariosAvailableInThisReq - ScenariosExecutedInThisReq
                     if RemaningScenario !=0:
                          Project.Variables.req_status='Fail'
        
                #The below code iterates through the XML and gets the TC count available for the mentioned Scenario
                for i in range(len(scen_ID_list)):

                            TestCasesInThisScenario = Testcases_list[i]
                              
                            #The below code iterates through the dict and gets the TC count executed for the mentioned Scenario
                            for key,val in Scen_dict.items():
                                         
                                        if key==scen_ID_list[i]:##Scen_ID
                                                TCExecutedInThisScen = len(Scen_dict.get(key))  
                                                                                  
                                                #Looks for any difference in TC count in XML vs actually executed
                                                if TCExecutedInThisScen!= TestCasesInThisScenario:
                                                      RemainingTC = TestCasesInThisScenario - TCExecutedInThisScen
                                                      if RemainingTC!=0:
                                                            Project.Variables.req_status='Fail'
          
          
                Header_HTMLReport()
                    
          
                for i in range(len(scen_ID_list)):
       
                            TestCasesInThisScenario = Testcases_list[i]
                              
                            #The below code iterates through the dict and gets the TC count executed for the mentioned Scenario
                            for key,val in Scen_dict.items():
                                         
                                        if key==scen_ID_list[i]:
                            
                                                TCExecutedInThisScen = len(Scen_dict.get(key)) 
                                                print_Scenario()
                                                for TC,Status in (Scen_dict.get(key)).items():
                                                      if (Scen_dict.get(key)).get(TC)==True:
                                                        Flag='Pass'
                                                      else:
                                                        Flag='Fail'
                                                      if TC in TC_ID_list:
                                                        j= TC_ID_list.index(TC)
                                                        Print_testCase(Flag)
                                                
                                                
                                                  
                                                
                                
                            #Looks for any difference in TC count in XML vs actually executed - Details not required
      #                      if TCExecutedInThisScen!= TestCasesInThisScenario:
      #                            RemainingTC = TestCasesInThisScenario - TCExecutedInThisScen
      #                            if RemainingTC!=0:
      #                                  aqFile.WriteToTextFile(HTML_Report_path, "<p style='font-size: 35px;color:#ff0000'>TC Remaining to be executed within this Scenario: "+(str(RemainingTC))+"</p>",aqFile.ctUTF8)
      #                                           
                            #Loop through TC LIsts to get the lists of TC not run
                                                               
                Print_Exception()                                      
                if ScenariosAvailableInThisReq != ScenariosExecutedInThisReq :
                   RemaningScenario= ScenariosAvailableInThisReq - ScenariosExecutedInThisReq
                   if RemaningScenario !=0:
                        aqFile.WriteToTextFile(HTML_Report_path, "<p style='font-size: 35px;color:#ff0000'>Scenario Remaining to be executed within this Requirement: "+(str(RemaningScenario))+"</p>",aqFile.ctUTF8)   
                                

        Print_Lists_of_TC_not_Executed()
        aqFile.WriteToTextFile(HTML_Report_path, "</body>",aqFile.ctUTF8)
        aqFile.WriteToTextFile(HTML_Report_path, "</html>",aqFile.ctUTF8)
        ConvertHTMLtoPDF()
        onStop_TestedItems()
  except Exception as e:
    save_exception(e)
    Print_Exception()
    Print_Lists_of_TC_not_Executed()
    aqFile.WriteToTextFile(HTML_Report_path, "</body>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</html>",aqFile.ctUTF8)
    ConvertHTMLtoPDF()
    onStop_TestedItems()
    
                
                  
def On_start_Req(ReqID):
   try:
        global Scen_dict,TestCasesInThisScenario,counter,screenshot_list,TC_executed_lists
        global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
        global Str_exception
        Project.Variables.ReqID=ReqID
        counter=0
        screenshot_list=[]
        TC_executed_lists =[]
        Str_exception=None
        Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list =xml(Project.Variables.ReqID)#ReqID
   except Exception as e:
        save_exception(e)
        Log.Message(str(e))
                  

  
def TC_validation_status(TC_ID,TC_Status,**result):
  try:
        global Scen_dict,TestCasesInThisScenario,index_of_TC
        ScenarioID = "_".join(TC_ID.split("_")[:4])
        Scen_dict.setdefault(ScenarioID,{}).update({TC_ID:TC_Status})
        take_Screenshot(TC_ID)

        if TC_ID in TC_ID_list:
             index_of_TC = TC_ID_list.index(TC_ID)
        else:
           Log.Message('Entered TC ID is invalid')
     
        for key,value in result.items():
            if key in TC_ActualPass_list[index_of_TC] :
              TC_ActualPass_list[index_of_TC] = TC_ActualPass_list[index_of_TC].format(**result)
              TC_ActualFail_list[index_of_TC] = TC_ActualFail_list[index_of_TC].format(**result)
              break
            else:
              Log.Message('Recheck the arguments entered')
        Replace_Actual_Fail(TC_ID,TC_Status)
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))
        

def Replace_Actual_Fail(TC_ID,TC_Status):
  global Str_exception
  if TC_Status ==False:
    if Str_exception !=None:
      exception_statement = Str_exception[0]+' : ' + Str_exception[1] +' : '+Str_exception[2]
      TC_ActualFail_list[index_of_TC]= exception_statement
  

                  
def onStop_TestedItems():
  try:
          global Req_ID,Req_Desc,Precondition,Testcases_list,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
          global Scen_dict,TestCasesInThisScenario,counter,screenshot_list,j,i,TC_executed_lists
          if Project.TestItems.Current!=None:
            Req_Level_Contents()
          
          scenarios_list=[]
          scen_ID_list=[]
          scen_Desc_list=[]
          Scen_dict={}
          SPreconditon_list =[]
          Testcases_list=[]
          TC_ID_list=[]
          TC_Description_list=[]
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
          TC_executed_lists =[] 
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))
                  
       
def Insert_Test_data(TC_ID,**kwargs):
  try:
     global Req_ID,Req_Desc,Precondition,Testcases_list,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
     global Scen_dict,TestCasesInThisScenario,counter,screenshot_list,j,i
     if TC_ID in TC_ID_list:
       index_of_TC = TC_ID_list.index(TC_ID)
     else:
       Log.Message('Entered TC ID is invalid')
     for key,value in kwargs.items():
        if key in TC_Action_list[index_of_TC] :
          TC_Action_list[index_of_TC] = TC_Action_list[index_of_TC].format(**kwargs)
   TC_Expected_list[index_of_TC] = TC_Expected_list[index_of_TC].format(**kwargs)
          break
        else:
          Log.Message('Recheck the arguments entered')
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))
   
     
def Insert_Test_Result(TC_ID,**result):
  try:
        global Req_ID,Req_Desc,Precondition,Testcases_list,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
        global Scen_dict,TestCasesInThisScenario,counter,screenshot_list,j,i
        if TC_ID in TC_ID_list:
             index_of_TC = TC_ID_list.index(TC_ID)
        else:
           Log.Message('Entered TC ID is invalid')
        for key,value in result.items():
            if key in TC_ActualPass_list[index_of_TC] :
              TC_ActualPass_list[index_of_TC] = TC_ActualPass_list[index_of_TC].format(**result)
              TC_ActualFail_list[index_of_TC] = TC_ActualFail_list[index_of_TC].format(**result)
              break
            else:
              Log.Message('Recheck the arguments entered')
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))
        
  
def Print_Exception():
  try:
          global Req_ID,Req_Desc,Precondition,Testcases_list,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
          global Scen_dict,TestCasesInThisScenario,counter,screenshot_list,j,i,HTML_Report_path
          global Str_exception
          if Str_exception != None:
  
                  aqFile.WriteToTextFile(HTML_Report_path, "<p style='font-size: 35px;color:#000000'> Exception occured: <mark><b>"+Str_exception[0]+ "</b></mark><br/>""The Type of exception is <mark><b>" +Str_exception[1]+"</b></mark> in File :<b><mark>" +Str_exception[2]+ "</b></mark>, at Line No: <mark><b>"+Str_exception[3]+"</b></mark></p>",aqFile.ctUTF8)
                  Project.Variables.req_status='Fail'
                  req_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y")
                  screenshot_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y_%I_%M_%S")
                  Foldername= Req_ID+'_'+ req_dateInfo
                  img_dir = Project.Path+'Screenshots\\'+Foldername
                  
                  screenshot_name = Req_ID+'_Exception_screenshot'+'_'+screenshot_dateInfo
                  if(Utilities.FileExists(img_dir+'\\'+screenshot_name+'.jpg') == False):
                                pic = Sys.Desktop
                                img_name= img_dir+'\\'+screenshot_name+'.jpg'
                                pic.Picture().SaveToFile(img_name)
                                Log.Message("screenshot File created :" +screenshot_name)
        #                        screenshot_list.append(screenshot_name)
        #                  j=j+1
                                aqFile.WriteToTextFile(HTML_Report_path, '<img src="%s" alt="screenshot" />'%(img_name),aqFile.ctUTF8)
                                Log.Message(screenshot_name+ " screenshot written to HTML" )
                                Str_exception=None
                  else:
                                Log.Message("Screenshot not created")
        #  else:
        #          pass
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))

  
def save_exception(e):
#  try:
      global Str_exception,exc_type,filename,exc_tb,tb_lineno
      import sys,os
      exc_type, exc_obj, exc_tb = sys.exc_info()
      fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
      filename=((fname.split(':'))[1])[:-1]
      
      Str_exception =[str(e),exc_type.__name__, filename, str(exc_tb.tb_lineno)]
      Log.Error(str(Str_exception))
#  except Exception as e:
#        save_exception(e)
#        Log.Message(str(e))
  

def Print_Lists_of_TC_not_Executed():
  try:
        global TC_executed_lists,Scen_dict,TC_ID_list,HTML_Report_path
  
        for TC_set in Scen_dict.values():
          for TC in TC_set.keys():
            TC_executed_lists.append(TC)
        if len(TC_executed_lists)!= len(TC_ID_list):  
              aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
              aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)  
              aqFile.WriteToTextFile(HTML_Report_path, "<table class='TCNotExecuted' style='align:center'>",aqFile.ctUTF8)
              aqFile.WriteToTextFile(HTML_Report_path, "<tr><td  colspan='2' style='text-align:center;font-weight: bold'>Details of TC Not Executed</td></tr> ",aqFile.ctUTF8)
              aqFile.WriteToTextFile(HTML_Report_path, "<tr><th width='25%'>Scenario Number</th> <th>TC Not Executed </th></tr>",aqFile.ctUTF8)
              for TC in TC_ID_list:
                if TC not in TC_executed_lists:
                    Scen_ID ="_".join(TC.split("_")[:4])
                    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td height='30'>"+(str(Scen_ID))+"</td><td>"+(str(TC))+"</td></tr>",aqFile.ctUTF8)
              aqFile.WriteToTextFile(HTML_Report_path, "</table>",aqFile.ctUTF8)
              aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
              aqFile.WriteToTextFile(HTML_Report_path, "<br/> ",aqFile.ctUTF8)
  except Exception as e:
        save_exception(e)
        Log.Message(str(e))
                                   

def HTML_Summary_Header():
 try:
     global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
     global j,i,Total_Test_CaseIn_Each_Scenario,LengthofScenarios
     global username,SSO,Serial,OS,CA,SV, HTML_Summary_Report_path
     
#      j=0
     username,SSO,Serial,OS,CA,SV = get_user_details("Daipayan","Machine1")
                    
     HTML_Summary_Report_path = HTML_Summary_Report_filepath +"Summary Report-"+aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%b_%Y_%H_%M_%S")+'.html'                  
     if(Utilities.FileExists(HTML_Summary_Report_path) == False):
         aqFile.Create(HTML_Summary_Report_path)
     Html_Styling('Summary')    
     Report_Header('Summary')
     Body_Req_details('Summary')
     Body_Environment_details('Summary')
 except Exception as e:
   save_exception(e)
 
 

def Module_level_Contents(*args):
 try:
   
   global TestGroup
   global HTML_Summary_Report_path                                                                  
   aqFile.WriteToTextFile(HTML_Summary_Report_path, "<table class='Mytbl'>",aqFile.ctUTF8)
   
   if len(args)!=0:
        aqFile.WriteToTextFile(HTML_Summary_Report_path, "<tr><td width='25%' colspan='5' style='text-align:center'><b>Requirement Not Aplicable for this Run</b> </td></tr> ",aqFile.ctUTF8) 
        aqFile.WriteToTextFile(HTML_Summary_Report_path, "<tr><td><b>Req ID</b></td><td colspan ='4'><b>Status</b></td></tr>",aqFile.ctUTF8)
   else:
       
       aqFile.WriteToTextFile(HTML_Summary_Report_path, "<tr><td width='25%' colspan='5' style='text-align:center'><b>"+Project.TestItems.Current.Parent.Name+"</b></td></tr> ",aqFile.ctUTF8)
       aqFile.WriteToTextFile(HTML_Summary_Report_path, "<tr><td><b>Req ID</b></td><td><b>Scenario ID</b></td><td><b>TC ID</b></td><td><b>Status</b></td><td><b>Link to PDF</b></td></tr>",aqFile.ctUTF8)         
 
 except Exception as e:
   save_exception(e)
     
def Req_Level_Contents(*args):
 try:
   global Req_ID,Req_Desc,Precondition,Testcases_list,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
   global TC_executed_lists,Scen_dict,TC_ID_list
   global HTML_Summary_Report_path,HTML_Report_path
   
   if len(args)!=0:
            Req_ID=args[0]
            status ='Not Applicable'
            aqFile.WriteToTextFile(HTML_Summary_Report_path, "<tr><td height='30'>"+Req_ID+"</td><td colspan='4'>"+status+"</td></tr>",aqFile.ctUTF8)  
    
   else:
           file_name = HTML_Report_path
           len_lists= len(TC_ID_list)
           for TC in TC_ID_list:
               Scenario_ID = "_".join(TC.split("_")[:4])
               if TC in TC_executed_lists:
                 if Scen_dict[Scenario_ID][TC]== True:
                   status = 'Pass'
                 elif Scen_dict[Scenario_ID][TC]== False:
                   status ='Fail'
               else:
                  status = 'Not Executed'
               aqFile.WriteToTextFile(HTML_Summary_Report_path, "<tr><td height='30'>"+Req_ID+"</td><td>"+Scenario_ID+"</td><td>"+TC+"</td><td>"+status+"</td><td><a href = '{file}'>View</a></td></tr>".format(file= file_name),aqFile.ctUTF8)  
           aqFile.WriteToTextFile(HTML_Summary_Report_path, "</table>",aqFile.ctUTF8)
           aqFile.WriteToTextFile(HTML_Summary_Report_path, "<br/> ",aqFile.ctUTF8)
           aqFile.WriteToTextFile(HTML_Summary_Report_path, "<br/> ",aqFile.ctUTF8)
           Log.Message('Requirement Summary Report printed to HTML')


#   else:
#           file_name = HTML_Report_path
#           
#           aqFile.WriteToTextFile(HTML_Summary_Report_path, '<tr><td height="30" rowspan = "{}">'+Req_ID+'</td>'.format(len(TC_ID_list)),aqFile.ctUTF8)  
#           for TC in TC_ID_list:
#               Scenario_ID = "_".join(TC.split("_")[:4])
#               if TC in TC_executed_lists:
#                 if Scen_dict[Scenario_ID][TC]== True:
#                   status = 'Pass'
#                 elif Scen_dict[Scenario_ID][TC]== False:
#                   status ='Fail'
#               else:
#                  status = 'Not Executed'
#               aqFile.WriteToTextFile(HTML_Summary_Report_path, "<td>"+Scenario_ID+"</td>",aqFile.ctUTF8)  
#               aqFile.WriteToTextFile(HTML_Summary_Report_path, "<td>"+TC+"</td>",aqFile.ctUTF8)  
#               aqFile.WriteToTextFile(HTML_Summary_Report_path, "<td>"+status+"</td>",aqFile.ctUTF8)
#               
#               if TC == TC_ID_list[-1]:   
#                  aqFile.WriteToTextFile(HTML_Summary_Report_path, "<td rowspan='{}'><a href = '{}'>View</a></td></tr>".format(len(TC_ID_list),file_name),aqFile.ctUTF8)
#               else:
#                  aqFile.WriteToTextFile(HTML_Summary_Report_path, "<td>""</td></tr>",aqFile.ctUTF8)
#           aqFile.WriteToTextFile(HTML_Summary_Report_path, "</table>",aqFile.ctUTF8)
#           aqFile.WriteToTextFile(HTML_Summary_Report_path, "<br/> ",aqFile.ctUTF8)
#           aqFile.WriteToTextFile(HTML_Summary_Report_path, "<br/> ",aqFile.ctUTF8)
#           Log.Message('Requirement Summary Report printed to HTML')


      
 except Exception as e:
   save_exception(e)




   
                   
def Closing_HTML_Not_selected_Summary_Report():
    
    aqFile.WriteToTextFile(HTML_Summary_Report_path, "</table>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Summary_Report_path, "<br/> ",aqFile.ctUTF8)