########license checker start####################copy pase the whole code below in main fuction

# hash input  is expiry date + content +salt concatenated in that order
#assumes current license key is saved in local license.key file is ascii 
#the below code has to be pasted as in in the main function .py container 
# calling of the function to be from the main function...caling sample given at the end
#last update by Shivaji Basu 122188 5 Aug 2021
def check_entitlement(content='',file='license.key',salt='synergy12'):
    
    import hashlib
    from datetime import datetime,timedelta
    

    def gethash(token:str='',salt:str=salt,length:int=12):
        try:
            hs=hashlib.sha256()
            token=token+salt
            hs.update(token.encode('utf-8'))
            return hs.hexdigest()[0:length]
        except:
            print("Error: hashing failed")

    def time_limit(serial:str):
        try:
            
            idate=datetime.now().date()
            for i in range(365*5):
                idate=idate+timedelta(1)
                sidate=idate.strftime('%Y-%m-%d')
                
                token=sidate+content
                skey=gethash(token=token,salt=salt)
                
                
                if skey.encode('utf-8')==serial.encode('utf-8'):
                    return idate
            return None
  
        except:
            raise
            print("Error in entitlement check")

    
    
    try:
        file1 =open(file,'r')
        serial=file1.read()
        
        file1.close()
        
    except:
        print("license key not found")
        quit()
        return False # in case quit() is overriden 

    try:
        last_date=time_limit(serial)
        if last_date!=None:
            current_date=datetime.now().date()
            if current_date>last_date:
                print("License expired.")
                quit()
                return False #incase quit is overriden
            else:
                print(f'License valid until {last_date}')
                return True
        else:
            
            print("license key not valid")
            quit()
            return False #incase quit is overriden
                
            print(f"License valid")# until {last_date}")
    except:
        
        print("Error checking license validity")
        raise
        quit()
        return False #incase quit is overriden





#################license checker end######################################



if __name__=="__main__":
    content="test content"# can be machine id, email address, schema file reference
    check_entitlement(content=content)


