from Report_Common_funcs import *
from HTML_Reporting_Dynamic import *

def test():
  try:
      global Str_exception
      ReqID= 'PS_APP_1525'
      On_start_Req(ReqID)

#      num1=5
#      num2=8
#      sum=num1+num2
#      Insert_Test_data('PS_APP_1525_SC01_TC01',num1 =5 ,num2=3)
#      if sum>10:
#        TC_validation_status('PS_APP_1525_SC01_TC01',True,sum=sum)
#        
#      else:
#        TC_validation_status('PS_APP_1525_SC01_TC01',False,sum=sum)
      

#      TC_validation_status('PS_APP_1525_SC01_TC02',True)
      
      
      TC_validation_status('PS_APP_1525_SC02_TC01',True,PID =123)
      Insert_Test_data('PS_APP_1525_SC02_TC01',address = 'blr')
      x=1
      y=3
      sum =x+y
      if (sum==6):
        TC_validation_status('PS_APP_1525_SC02_TC02',True,sum=sum,number=12)
      else:
        TC_validation_status('PS_APP_1525_SC02_TC02',False,sum=sum,number=12)
      
#      TC_validation_status('PS_APP_1525_SC02_TC03',True)
#      Insert_Test_data('PS_APP_1525_SC02_TC03',LastName = 'Ganga',FirstName = 'Nimith',PID =234)
            
  except Exception as e:
    Log.Error(str(e)) 
    #New Line of Code   
    Str_exception = str(e)
    save_exception(Str_exception)
      
  finally:
    HTML_Report_print()
    
    

  

  
         
