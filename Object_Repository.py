
from common_variables import *
#-------------------------------------------------------------------------------------------------------
class Set_up():
    setup_btn =[['WndClass','WndCaption'],['Button','Setup']]
    
class Set_up_window(): 
    setup_dialog = [['ObjectType','Caption'],['Dialog','Setup']]
    tab_Current_Telemetry = [['ObjectType','ObjectIdentifier'],['PageTab','Current Listings']]
   
    
class Central_defaults(): 
  unitname = [['Name','ObjectType','Visible'],['Window("ComboBox", "", 1)','ComboBox','True']]

    

class User_Setup(): #User_Setup.User_Setup_screen
    User_Setup_screen = [['ObjectType','ObjectIdentifier'],['PropertyPage','User Setup']]
    user_passwordbox = [['ObjectType','ObjectIdentifier'],['Edit','Password:']]
    
    
    
