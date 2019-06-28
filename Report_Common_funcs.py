import re
from common_variables import *
import os
import xml.etree.ElementTree as ET
global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
#-------------------------------------------------------------------------------
# REPORTING COMMON FUNCTIONS
#-------------------------------------------------------------------------------
#Purpose: Get the feature file in text prescribed format.
#Author: Daipayan               Date:25/2/2019 
#Input: Filename refers to the entire FilePath
#-------------------------------------------------------------------------------


#def getFile(Filename):
#      File=aqFile.ReadWholeTextFile(Filename,aqFile.ctUnicode)
#      return File
      
#-------------------------------------------------------------------------------
#Purpose: Get the feature file in text prescribed format.
#Author: Daipayan               Date:25/2/2019 
#Input: File refers to the entire File contents in str derived from getFile() && var_name is the specific property to derive.
#-------------------------------------------------------------------------------


#def readFile(File,var_name):
# # File=aqFile.ReadWholeTextFile(Filename,aqFile.ctUnicode)
#  contents_lists= File.split('\r\n')
#  for each_line in contents_lists:
#    spec_item = each_line.split('=')
#    if spec_item[0]==var_name:
#      return spec_item[1]
   
#-------------------------------------------------------------------------------
#Purpose: Get and read the xml based feature file in prescribed format.
#Author: Daipayan               Date:5/3/2019 
#Input: ReqID: Requirement ID
#       ScenarioID: Scenario ID as per the feature file
#       TestID: Test ID as per the feature file
#       Flag : Y or N
#-------------------------------------------------------------------------------


def read_req_xml(ReqID,ScenarioID,TestID,Flag):
    try:
        req_dir =Project.Path + '\\FeatureFiles'
        for root,dirs,files in os.walk(req_dir):
            for req_name in files:
                if req_name.startswith(ReqID):
                    file_name = os.path.join(root,req_name)
                    if aqFile.Exists(file_name):
                          Log.Message('File exists')
                      
                          #Parsing XML starts here
                          tree = ET.parse(file_name)
                          root = tree.getroot()
                          for childnodes in root:
                               # get Pre condition
                                if childnodes.tag == 'PreCondition':
                                        Log.Message(childnodes.text)
                                if childnodes.tag == 'Req':
                                        Log.Message(childnodes.get('ID'))
                                        Log.Message(childnodes.get('Description'))
                                       # Log.Message scenario ID and description
                                        for child in childnodes:  
                                            if child.tag == 'Scenario' and child.get('Automate') =='Yes' and child.get('ID') == ScenarioID :
                                                    Log.Message(child.get('ID'))
                                                    Log.Message(child.get('Description'))
                                                    for test in child:
                                                      if test.tag == 'TestCase' and test.get('ID')==TestID:
                                                            Log.Message(test.find('Action').text)
                                                            Log.Message(test.find('Expected').text)
                                                            if (Flag == 'Y'):
                                                              if test.find('ActualPass').text != None:
                                                                Log.Message(test.find('ActualPass').text)
                                                              else:
                                                                Log.Message('Pass result is empty')
                                                            else:
                                                                Log.Message(test.find('ActualFail').text)
                                                    return

                                            else:
                                              Log.Message('Scenario input is invalid. Please recheck data')
                                              return
                                              
                                                                
                    elif ET.parse(file_name).parseError.errorCode != 0:
                          s = "Reason:\t" + ET.parse(file_name).parseError.reason + "\n" + "Line:\t" + aqConvert.VarToStr(ET.parse(file_name).parseError.line) + "\n" + "Pos:\t" + aqConvert.VarToStr(ET.parse(file_name).parseError.linePos) + "\n" + "Source:\t" + ET.parse(file_name).parseError.srcText
                          # Post an error to the log and exit
                          Log.Error("Cannot parse the document.", s)
                          
                    else:
                            Log.Message('File doesnot exists')
                       


    except Exception as e:
        Log.Message(str(e))

#-------------------------------------------------------------------------------
#Purpose: Gets the current Unit name.
#Author: Daipayan               Date:5/3/2019 
#Output: str 
#-------------------------------------------------------------------------------
    
import os, inspect
import re

def Get_current_Filename():
    sPath = inspect.getfile(inspect.currentframe())
#    sPath = inspect.stack()[1]
    str_file = (re.sub('[(){}<>]','',sPath))
    filename = str_file.strip('aq:')
    sName = aqFileSystem.GetFileName(filename)
    return sName

#-------------------------------------------------------------------------------
#Purpose: creates an ID(to be appended to HTMl) based on datetime to identify reports uniquely..
#Author: Daipayan               Date:7/3/2019 
#Output: str 
#-------------------------------------------------------------------------------

def Get_unique_report_id():
  CurrentDate = aqDateTime.Today()
  CurrentTime = aqDateTime.Time()
  # Return the parts of the current date&time
  Year = aqDateTime.GetYear(CurrentDate)
  Month = aqDateTime.GetMonth(CurrentDate)
  Day = aqDateTime.GetDay(CurrentDate)
  Hours = aqDateTime.GetHours(CurrentTime)
  Minutes = aqDateTime.GetMinutes(CurrentTime)
  if(Month<10):
    Month = "0" + str(Month)
  if(Day<10):
    Day = "0" + str(Day)
  if(Hours<10):
    Hours = "0" + str(Hours)
  if(Minutes<10):
    Minutes = "0" + str(Minutes)
  URID = Get_current_Filename()+"_" + str(Month)+""+str(Day)+""+str(Hours)+""+str(Minutes)
  return URID
  
#-------------------------------------------------------------------------------
#Purpose: Read data from xml and creates various lists . To be used for HTML printing.
#Author: Daipayan               Date:11/3/2019 
#Output: LIsts 
#Latest Update: Included TC_Descrtiption in xml and also the reurn value , and global variable here
#-------------------------------------------------------------------------------


import os
import xml.etree.ElementTree as ET
import re
          
def xml(ReqID):
  global Req_ID,Req_Desc,Precondition,total_scenarios,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
  try:
      scenarios_list=[]
      scen_ID_list=[]
      scen_Desc_list=[]
      SPreconditon_list =[]
      Testcases_list=[]
      TC_ID_list=[]
      TC_Description_list=[]
      TC_Action_list=[]
      TC_Expected_list =[]
      TC_ActualPass_list=[]
      TC_ActualFail_list =[]
      
      req_dir =Project.Path + '\\FeatureFiles'
      for root,dirs,files in os.walk(req_dir):
            for req_name in files:
             if req_name.startswith(ReqID):
                 file_name = os.path.join(root,req_name)
                 if aqFile.Exists(file_name):
                      Log.Message('File exists')
                      #Parsing XML starts here
                      tree = ET.parse(file_name)
                      root = tree.getroot()
                      
                      for childnodes in root:
                               # get Pre condition
                        if childnodes.tag == 'PreCondition':
                          Precondition = (childnodes.text)
                          
                        if childnodes.tag == 'Req':
                          Req_ID = (childnodes.get('ID'))
                          Req_Desc = (childnodes.get('Description'))


                      for childnodes in root.findall('Req'):
                        scenarios= childnodes.findall('.//Scenario[@Automate="Yes"]')
                        
                            
                            #Gives the total no of scenarios inside the Req XML
                        total_scenarios = (len(scenarios))
                        
                        Log.Message("Total Test scenarios in present in this requirement is:" +str(total_scenarios))
                        for each_scenario in scenarios:
                              if each_scenario.get('Automate')=='Yes':
                                  scen_ID = (each_scenario.get('ID'))
                                  scen_Desc =   (each_scenario.get('Description'))
                                  SPreconditon = each_scenario.find('Scenario_PreCondition').text
                                                                   
                                  scen_ID_list.append(scen_ID)
                                  scen_Desc_list.append(scen_Desc)
                                  if each_scenario.find('Scenario_PreCondition').text!=None:
                                    SPreconditon_list.append(SPreconditon)
                                  else:
                                    SPreconditon_list.append(None)
                                  Testcases = each_scenario.findall('TestCase')
                                  total_Test=len(Testcases)
                                  Log.Message("Total Test Cases in present in Scenario {} is: {}" .format(scen_ID,total_Test))
                                  
                                  Testcases_list.append(total_Test) #total_Test gives the total Test cases under each scenario
      
                                  for each_test in Testcases:
                                      TC_ID=(each_test.get('ID'))
                                      TC_Description= (each_test.find('TC_Description').text)
                                      TC_Action= (each_test.find('Action').text)
                                      TC_Expected =(each_test.find('Expected').text)
                                      TC_ActualPass =(each_test.find('ActualPass').text)
                                      TC_ActualFail = (each_test.find('ActualFail').text)
                                                                      
                                      TC_ID_list.append(TC_ID)
                                      TC_Description_list.append(TC_Description) 
                                      TC_Action_list.append(TC_Action)
                                      TC_Expected_list.append(TC_Expected)
                                      TC_ActualPass_list.append(TC_ActualPass)
                                      TC_ActualFail_list.append(TC_ActualFail)
            return Req_ID,Req_Desc,Precondition,Testcases_list,scenarios, scen_ID_list,scen_Desc_list,SPreconditon_list , Testcases_list,TC_ID_list,TC_Description_list,TC_Action_list,TC_Expected_list,TC_ActualPass_list,TC_ActualFail_list
                                                           
  except Exception as e:
      Log.Message(str(e))

  else:
       Log.Message('File doesnot exists')
       
#-------------------------------------------------------------------------------
#Purpose: Gets text the file to be read.
#Author: Daipayan               Date:7/3/2019 
#Output: str 
#-------------------------------------------------------------------------------       

def getFile(Filename):
      File=aqFile.ReadWholeTextFile(Filename,aqFile.ctUTF8)
      return File

#-------------------------------------------------------------------------------
#Purpose: Reads the text the file .
#Author: Daipayan               Date:7/3/2019 
#Output: str 
#------------------------------------------------------------------------------- 

def readFile(File,var_name,prop_name):

  contents_lists= File.split('\r\n')
  for each_line in contents_lists:
    spec_item = each_line.split('=')
    if spec_item[0]==var_name:
      dict = spec_item[1].split(',')
      for i in range(len(dict)):
        str_val = dict[i].split(':')
        if str_val[0]==prop_name:
          return str_val[1]
