from common_variables import *
from Object_Repository import *
from HTML_report_funcs import *
import datetime
##------------------------------------------------------------------------------
##Purpose: identifies an Object using the passed arguments and returns the object.
##Author: Daipayan            Date:5/2/2019 
##Parent to be selected from common_variables
##ele_name to be selected from Object_Repository
##-------------------------------------------------------------------------------
#def find_obj(parent,ele_name):
#  propArr= ele_name[0]
#  propVal = ele_name[1]
#  obj = parent.find(propArr, propVal, 30)
#  return obj

  
#------------------------------------------------------------------------------
#Purpose: identifies an Object using the passed arguments and returns the object.
#Author: Daipayan            Date:5/2/2019 
#Parent to be selected from common_variables
#ele_name to be selected from Object_Repository
#-------------------------------------------------------------------------------

def find_obj(parent,ele_name): 
  try:
    propArr=ele_name[0]
    propVal=ele_name[1]
    if len(propArr)==len(propVal):
      obj = parent.find(propArr,propVal,100)
      Delay(med_delay)
      if obj.Exists:
  #    if obj.WaitProperty("Exists",True,1500):
        Log.Message(obj.name+ " obj exists")
        return obj
      elif aqString.Find("CARESCAPE Central Station",propVal[1], False, False) != -1:
        return False
      else:
        Log.Error(ele_name+ " obj does not exists")
        return False
    else:
      Log.Error("Input properites are wrong. Recheck properites passed")
      return False
  except Exception as e:
    Log.Message(str(e))
    return False
    

#------------------------------------------------------------------------------
#Purpose: identifies an Object using the passed arguments for its parent and then returns the object.
#Author: Daipayan            Date:21/2/2019 
#Parent to be selected from common_variables
#add_property is the additonal prop argument that needs to be passed
#add_propVal is the additonal prop value argument that needs to be passed
#ele_name to be selected from Object_Repository
#-------------------------------------------------------------------------------

#def find_obj_custom(parent,ele_name,add_property,add_propVal): 




#  try:
#    propArr=ele_name[0]
#    propVal=ele_name[1]
#    if len(propArr)==len(propVal):
#      propArr.append(add_property)
#      propVal.append(add_propVal)
#      obj = parent.findChild(propArr,propVal,100)
#      Delay(med_delay)
##      if obj.Exists:
#      if obj.WaitProperty("Exists",True,1500):
#        Log.Message(obj.name+ " obj exists")
#        return obj
#      else:
#        Log.Error(ele_name+ " obj does not exists")
#        return false
#    else:
#      Log.Error("Input properites are wrong. Recheck properites passed")
#      
#  except Exception as e:
#    Log.Message(str(e))
    
#-------------------------------------------------------------------------------    
#Purpose: Clicks on the Object passed as arguments.
#Author: Daipayan            Date:5/2/2019 
#Parent to be selected from common_variables
#ele_name to be selected from Object_Repository
#Modified by: Daipayan - Placed enabled property inside if condition once it exists to avoid run time error
#-------------------------------------------------------------------------------
def click(obj):
 try:
   
  if (obj.Exists== True):
    obj.Refresh()
    if (obj.Enabled== True):
      obj.Click()
      return True
    else:
      Log.Message(obj+ ' does exists but not enablled to click')
      return False
  else:
      Log.Message(obj+ ' does not exists to click')
      return False
 except Exception as e:
    Log.Message(str(e))
    return False

#---------------------------------------------------------------------------------------------------        
#Purpose: Verifies the screen the user is in or verifies an object available in the intended screen
#Author: Daipayan            Date:6/2/2019
#Parent to be selected from common_variables
#ele_name is the property of the element/screen we intend to see is to be selected from Object_Repository
#exp_property is the property of the obj created for  ele_name. Example: Visble/Exists/Enabled
#--------------------------------------------------------------------------------------------------- 

def verify_obj_screen(obj_screen,exp_property):   
  if obj_screen!=False or obj_screen!= None:
    if obj_screen.WaitProperty(exp_property,True,med_delay):
      Log.Message('{} is present and in  {} state'.format(obj_screen.name,exp_property))
      return True
    else:
      Log.Message('{} is present but not in  {} state'.format(obj_screen.name,exp_property))
      return False
      
  else:
    return False
    
    
#-------------------------------------------------------------------------------    
#Purpose: Takes Screenshots.
#Author: Daipayan              Date:6/2/2019 
#Parent to be selected from common_variables
#-------------------------------------------------------------------------------
#
#def take_screenshot():
#  pic = Sys.Desktop.ActiveWindow()
#  pic.Picture().SaveToFile(Project.Path+'\\Screenshots\\screenshot_filename.jpg')
  
  
#-------------------------------------------------------------------------------    
#Purpose: Selecting from Dropdown
#Author: Daipayan            Date:7/2/2019 
#Parent to be selected from common_variables
#Obj should the dropdown element
#-------------------------------------------------------------------------------

def select_dropdown(obj,item_select):
  obj.ClickItem(item_select)
  if obj.Value==item_select:
    Log.Message(item_select+ "is selected for the dropdown " +select_dropdown.__name__)
  else:
    return False

    
    
#-------------------------------------------------------------------------------    
#Purpose: Selecting from Dropdown using specific Parent
#Author: Daipayan            Date:21/2/2019 
#Parent for this function needs to created and then used for finding object
#Obj should the dropdown element
#-------------------------------------------------------------------------------

def select_dropdown_using_parent(parent_obj,ele_name,item_select):
  obj = find_obj(parent_obj,ele_name)
#  obj.ClickItem(item_select)
  obj.Keys(item_select)
  if obj.parent.Value==item_select:
    Log.Message(item_select+ "is selected for the dropdown " +select_dropdown_using_parent.__name__)
  else:
    Log.Message(item_select+ " is not presnt using " +select_dropdown_using_parent.__name__+ "function")

    
##-------------------------------------------------------------------------------    
##Purpose: select dropdown for a special_case 
##Eample: selecting values from ComboBox peresent while admitting patient
##Author: Daipayan            Date:21/2/2019 
##Parent for this function needs to created and then used for finding object
##Obj should the dropdown element
##-------------------------------------------------------------------------------
#
#def select_dropdown_special_case(parent_obj,ele_name,item_select):
#  obj = find_obj(parent_obj,ele_name)
#  obj.ClickItem(item_select)
#  obj.Keys(item_select)
#  if obj.wText==item_select:
#    Log.Message(item_select+ "is selected for the dropdown " +select_dropdown_using_parent.__name__)
#  else:
#    Log.Message(item_select+ " is not presnt using " +select_dropdown_using_parent.__name__+ "function")

#-------------------------------------------------------------------------------  
#Purpose: Selecting from Dynamic Dropdown
#Author: Daipayan            Date:7/2/2019 
#Parent to be selected from common_variables
#Obj should the dropdown element
#-------------------------------------------------------------------------------

def dynamic_dropdown(obj,item_select):
  try:
   for val in obj.wItemList:
    if val == item_select:
      obj.ClickItem(item_select)
    else:
      Log.Message(item_select+ ' not avilable in ' +obj.name+ 'dropdownn')
  except Exception as e:
    Log.Message(str(e))
    
#-------------------------------------------------------------------------------------------
#Function Name :  ReadFileLines
#Purpose: To read required input data from text file
#Parameters: Input data FilePath, Input data Field Name 
#ReturnValue: Input data field value
#Author: Daipayan
#Eg: Utility_Functions.ReadFileLines(AFileName,'Sc1_Tc001') - File path and 'Sc1_Tc001'- test case name in i/p data file
#-------------------------------------------------------------------------------------------
def ReadFileLines(AFileName,Test_Data_Location):
#def ReadFileLines(AFileName,Input_Data_Field_Name,Test_Data_Location):
  Data_Found_Flag = 0
  F = aqFile.OpenTextFile(AFileName, aqFile.faRead, aqFile.ctANSI)
  F.Cursor = 0
  while not F.IsEndOfFile():
    s = F.ReadLine()
    if s.find(Test_Data_Location)!=-1:
      Input_Data = s.replace(Test_Data_Location,'')
      Input_Data_Dict = eval(Input_Data)
      return Input_Data_Dict
#      Input_Data_Dict = eval(s.replace(Test_Data_Location,''))
      #Log.message(Input_Data_Dict['TC_Description'])
      break
      Data_Found_Flag = 1
  if Data_Found_Flag != 1:
    Log.Message('Input Data is not available for the test case:{}'.format(Test_Data_Location))
  F.Close()
  


#-------------------------------------------------------------------------------
#Purpose: To generate HTML report Skeleton    
#Author: Daipayan  
#Date:22/3/2019  
#Last Upadted Date: 25/3/2019        
#Last Updated By:Daipayan 
#-------------------------------------------------------------------------------
def Create_Report(ReqID):
  Header_HTMLReport(ReqID)#Bl
  print_Scenario()
  
  
#-------------------------------------------------------------------------------
#Purpose: To Fill Text Box  
#Author: Daipayan  
#Date:29  /3/2019
#Last Upadted Date: 29/3/2019        
#Last Updated By:Daipayan 
#-------------------------------------------------------------------------------
def Enter_Text(Targetted_Obj, Value_To_Type):
  if Targetted_Obj.Exists and Targetted_Obj.Enabled:
    Targetted_Obj.SetText('')
    Targetted_Obj.SetText(Value_To_Type)
    return True
  else:
    Log.Message('Targetted Edit Box field is not present/Disabled on the screen')
    return False


  

  
 
