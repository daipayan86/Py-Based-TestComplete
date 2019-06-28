from common_variables import *
from Report_Common_funcs import *
import os
import xml.etree.ElementTree as ET
import re
from Report_Common_funcs import *
#ReqID
global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
global j,i,x,l
global username,SSO,Serial,OS,CA
j=0
i=0
#-------------------------------------------------------------------------------
#Purpose: Returns an Unique ID for using it further`
#Author: Daipayan                     Date:28/2/2019 
#------------------------------------------------------------------------------- 

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


#def Get_current_project_suite_name():
#
#  projectSuiteName = aqFileSystem.GetFileName(ProjectSuite.FileName)
#  projectSuiteName = projectSuiteName.substring(0, projectSuiteName.length - 4);
#  return projectSuiteName    
#
#
#def Get_current_project_name():
#
#  projectName = aqFileSystem.GetFileName(Project.FileName)
#  projectName = projectName.substring(0, projectName.length - 4);
#  return projectName    

#-------------------------------------------------------------------------------
#Purpose: Returns the Current Tested Items name`
#Author: Daipayan                     Date:28/2/2019 
#-------------------------------------------------------------------------------


def Get_current_tested_req_name():

  try:
    TestItems = Project.TestItems;
    Log.Message("The " + TestItems.Current.Name + " test item is currently running.",TestItems.Current.Description);
    return TestItems.Current.Name;
  
  except Exception as e:
    Log.Message(str(e))


#-------------------------------------------------------------------------------
#Purpose: Returns the Current Tested Items Descriptions`
#Author: Daipayan                     Date:28/2/2019 
#-------------------------------------------------------------------------------

def Get_current_testitem_desc():

  try:
    TestItems = Project.TestItems;
    Log.Message("The " + TestItems.Current.Name + " test item is currently running.",TestItems.Current.Description);
      
    return   TestItems.Current.Description;
  except Exception as e:
    Log.Message(str(e))
    
#-------------------------------------------------------------------------------
#Purpose: Returns the users and machine details stored in the TXT file in the prescribed format`
#Author: Daipayan                     Date:18/3/2019 
#-------------------------------------------------------------------------------


def get_user_details(userDetail,machineDetail):
  from Report_Common_funcs import readFile
  Filestr=getFile(User_details)
#  username = readFile(Filestr,"Daipayan","Name")
#  SSO= readFile(Filestr,"Daipayan","SSOID")
#  Serial = readFile(Filestr,"Machine1","Serial Number")
#  OS = readFile(Filestr,"Machine1","OS")
#  CA = readFile(Filestr,"Daipayan","CA")

  username = readFile(Filestr,userDetail,"Name")
  SSO= readFile(Filestr,userDetail,"SSOID")
  Serial = readFile(Filestr,machineDetail,"Serial Number")
  OS = readFile(Filestr,machineDetail,"OS")
  CA = readFile(Filestr,machineDetail,"CA")
  SV= readFile(Filestr,machineDetail,"SV")
  return username,SSO,Serial,OS,CA,SV

  
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: creates the HTML file header and other static deatils in the prescribed format. It takes Req ID as the input and parses the XML feature file with same name.
#Author: Daipayan                     Date:13/3/2019 
#---------------------------------------------------------------------------------------------------------------------------------------------------------
  

def Header_HTMLReport(ReqID):
      global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
      global j,i,x,l
      global username,SSO,Serial,OS,CA,SV
      j=0
      username,SSO,Serial,OS,CA,SV = get_user_details("Daipayan","Machine1")
      Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list =xml(ReqID)#ReqID
      l=len(scenarios)
      if len(scenarios)==len(Testcases_list):
        
      
        if(Utilities.FileExists(HTML_Report_path) == False):
            aqFile.Create(HTML_Report_path)
        Html_Styling()    
        Report_Header()
        Body_Req_details()
        Body_Environment_details()
        for i in range(l):
            x=Testcases_list[i]
            return
          
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Prints the scenario in the HTML file and appends HTML file header in the prescribed format. 
#It reads the value from the XML file taken Header_HTMLReport and appends it.
#Author: Daipayan                     Date:13/3/2019 
#--------------------------------------------------------------------------------------------------------------------------------------------------------

def print_Scenario():   
  global i,l,x
  if scenarios[i].get('Automate')=='Yes':
      Body_Scenario_details()
      i=i+1
      return

#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Prints the Test Cases in the HTML file and appends HTML file header in the prescribed format. 
#It reads the value from the XML file taken Header_HTMLReport and appends it.
#Author: Daipayan                     Date:13/3/2019 
#--------------------------------------------------------------------------------------------------------------------------------------------------------
                     

def Print_testCase(Flag):
  global x,j
  for tc in range(x):
    Body_TC_details(Flag)
    j=j+1
    return
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the CSS part of the HTML file. 
#Author: Daipayan                     Date:11/3/2019 
#--------------------------------------------------------------------------------------------------------------------------------------------------------
  
                                                                          
def Html_Styling():
          
          oFile = aqFile.OpenTextFile(HTML_Report_path, aqFile.faWrite, aqFile.ctUTF8, True)
          oFile.WriteLine("<!DOCTYPE html> ")
          oFile.WriteLine("<html>")
 #        oFile.WriteLine("<head><script src='Js/jquery-2.1.4.min.js'></script>\n<script src='Js/toggleImage.js' type='text/javascript'></script>")
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
          oFile.WriteLine("img[alt='screenshot']{width:100%;height:100%}")
          oFile.WriteLine("</style> ")
          oFile.WriteLine("</head> ")
          oFile.WriteLine("<body> ")
          oFile.WriteLine("<div style='border: 1px solid; position:fixed; top:50px; left:50px; display: none' class='arrow_image'><img src='' width='900px'; height= '550px';></div>")
          oFile.Close()
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the REport Header and prints to the HTML file.  
#Author: Daipayan                     Date:12/3/2019 
#--------------------------------------------------------------------------------------------------------------------------------------------------------

                                                                     
def Report_Header():
    global username,SSO,Serial,OS,CA,SV
    aqFile.WriteToTextFile(HTML_Report_path, "<div class='header'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, '<img src="C:\\Work\\CSCS\\CSCS_Automation\\CSCS_Automation\\InputData\\Logo\\GE_Logo.jpg" alt="logo" />',aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<h1>" +Req_ID+  ": Execution Report</h1>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</div><br>",aqFile.ctUTF8)
    
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the Requirement details and prints to the HTML file. 
#Author: Daipayan                     Date:12/3/2019 
#--------------------------------------------------------------------------------------------------------------------------------------------------------

       
def Body_Req_details():
    global username,SSO,Serial,OS,CA,SV
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='Mytbl' >",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement Name</td><td width='55%'>" +Req_Desc+ "</td><td style='text-align:center'>Status</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='20%'>Requirement PreCondition</td><td width='55%'>" +Precondition+ "<td rowspan ='5' style='text-align:center'></td></tr> ",aqFile.ctUTF8)                                                                      
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>SSO ID</td><td>" +SSO+ "</td>" + "</tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Executed By</td><td>" +username+ "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Date & Time Executed</td><td>" + str(aqDateTime.Now()) + "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the UUT details and prints to the HTML file. Takes input the User Details and machine Details txt file.
#Author: Daipayan                     Date:12/3/2019 
#--------------------------------------------------------------------------------------------------------------------------------------------------------

def Body_Environment_details():
    global username,SSO,Serial,OS,CA,SV                                                                       
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='Mytbl'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='25%' colspan='4' style='text-align:center'>UUT Details</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Serial Number</td><td>CA</td><td> OS:</td><td> SV:</td></tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td height='30'>"+Serial+"</td><td>"+CA+"</td><td>"+OS+" </td><td>"+SV+" </td></tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</table>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the Scenario table to be used for printing to HTML .Takes input the Feature XML File.
#Author: Daipayan                     Date:12/3/2019 
#--------------------------------------------------------------------------------------------------------------------------------------------------------

                                                                    
def Body_Scenario_details():
    global i
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='names' style='font-weight: bold'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td width='25%'>Scenario ID</td><td>" +scen_ID_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Scenario Name</td><td>" +scen_Desc_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    Body_Scenario_Precondition()
    aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the Scenario precondition and appends to the Scenario table in HTML if only available .Takes input the Feature XML File.
#Author: Daipayan                     Date:12/3/2019 
#------------------------------------------------------------------------------------------------------------------------------------------------------

    
def Body_Scenario_Precondition():
    global i
    if SPreconditon_list[i]!= None:
      aqFile.WriteToTextFile(HTML_Report_path, "<tr><td>Scenario Precondtion</td><td>" +SPreconditon_list[i]+ "</td></tr> ",aqFile.ctUTF8)
    else:
      pass
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the Test Case table in HTML .Takes input the Feature XML File.
#Author: Daipayan                     Date:12/3/2019 
#------------------------------------------------------------------------------------------------------------------------------------------------------

                                                                    
def Body_TC_details(Flag):
    global j,i
    aqFile.WriteToTextFile(HTML_Report_path, "<table class='names' style='font-weight: normal'>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='8%'>TestCase ID</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>Action</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>Expected result</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>Actual result</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='10%'>Status</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<tr> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='8%'>"+TC_ID_list[j]+"</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>"+TC_Action_list[j]+"</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>"+TC_Expected_list[j]+"</th>  ",aqFile.ctUTF8)
    print_Actual(Flag)
    aqFile.WriteToTextFile(HTML_Report_path, "</tr>",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "</table> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<br> ",aqFile.ctUTF8)
    take_Screenshot()
    
#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Defines the Actual result and pushes to the TC_Details_Function to be printed in HTML .Takes as input the Feature XML File.
#Author: Daipayan                     Date:12/3/2019 
#------------------------------------------------------------------------------------------------------------------------------------------------------

def print_Actual(Flag):
  global j,i
  if (Flag=='Pass'):
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>"+TC_ActualPass_list[j]+"</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='10%' style = 'background-color: #32CD32'>Pass</th>  ",aqFile.ctUTF8)
  elif Flag=='Fail':
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>"+TC_ActualFail_list[j]+"</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='10%' style = 'background-color: #FF0000'>Fail</th>  ",aqFile.ctUTF8)
  else:
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='30%'>""</th>  ",aqFile.ctUTF8)
    aqFile.WriteToTextFile(HTML_Report_path, "<th width='10%'>Not Executed</th>  ",aqFile.ctUTF8)
    

#---------------------------------------------------------------------------------------------------------------------------------------------------------
#Purpose: Takes sceenshot and stores in a folder .Takes the screenshot and post it to HTML . 
#Author: Daipayan                     Date:12/3/2019 
#------------------------------------------------------------------------------------------------------------------------------------------------------


    
def take_Screenshot():
  global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
  try:
    ReqID =Req_ID
    global j,i
    req_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y")
    screenshot_dateInfo = aqConvert.DateTimeToFormatStr(aqDateTime.Now(),"%d_%m_%Y_%I_%M_%S")
    Foldername= ReqID+'_'+ req_dateInfo
    img_dir = Project.Path+'Screenshots\\'+Foldername
    screenshot_filename = ReqID+TC_ID_list[j]+screenshot_dateInfo
    if i==1:
      if(Utilities.DirectoryExists(img_dir)==True):
        Log.Message("Screenshot_Req Folder with same name already exists :" +Foldername )
      else:
        aqFileSystem.CreateFolder(img_dir)
        Log.Message("Screenshot_Req Folder created" +Foldername)
    if(Utilities.FileExists(img_dir+'\\'+screenshot_filename+'.jpg') == False):
          pic = Sys.Desktop
          img_name= img_dir+'\\'+screenshot_filename+'.jpg'
          pic.Picture().SaveToFile(img_name)
          Log.Message("screenshot File created : " +screenshot_filename+".jpg")
          aqFile.WriteToTextFile(HTML_Report_path, '<img src="%s" alt="screenshot" />'%(img_name),aqFile.ctUTF8)

    else:
         Log.Message("Screenshot not created")
  except Exception as e:
    Log.Message(str(e))