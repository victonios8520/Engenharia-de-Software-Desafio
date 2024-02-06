import os.path


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1gE0t5Dzp7ZyGCXPKmbHJU_koVZCCP9Gm3bT0ZVQRkDU" #sheet code
SAMPLE_RANGE_NAME = "Pagina1!A1:H27"#cell range


def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # login
  #api authentication
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())
      
  print("*"*50)
      
  print("Authentication was successful")

  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    
    print("*"*50)
    
    print("The reading was successful")
    
    print(f"The total number of absences allowed is {getNumberAbsencesCommitted(result['values'])}")
    
    print("*"*50)
    
    #This function update the sheet and returns the value of the new table
    
    try :
      #the function below returns the list with the new values ​​containing the final grade and the final mention
      
      readyResult = getStudentValue(result['values'],getNumberAbsencesCommitted(result['values']))
      
      # Write the new information on the sheet
  
      print("*"*50)
      
      print("Final mention of all students was successfully calculated")
      
      print("*"*50)
      
      service = build("sheets", "v4", credentials=creds)

      sheet = service.spreadsheets()
      
      body = {"values": readyResult}
      
      result = (
          sheet.values()
          .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range= "A1",valueInputOption= "USER_ENTERED",body = body)
          .execute()
      )
    except ValueError:
      print("Some value was entered incorrectly, please check the notes and absences")
    except:
      print("There was a formatting error, check the spreadsheet and correct the values ​​and try again")
    else:
       print("The dashboard has been updated successfully")
    
  except HttpError as err:
    print(err)
    
  
    
def getNumberAbsencesCommitted(dictionary):
  #This function takes the minimum number of absences that the student can have based on the value of classes in the sheets
  
  numberClass = 0
  
  for i in dictionary:
    for j in i:
      if('Total de aulas' in j):
        
        numberClass = j #get the j value
        
        numberClass = numberClass.split(":") # separates the string from the integer
        
        numberClass.pop(0) #remove the string
        
        numberClass = int(numberClass[0]) #transforms the number into an integer
        
        numberClass *= 0.25 #only get 25 percent
        
  return numberClass # retun the 25 percent

def getStudentValue(student,numberClass):
  
  grade = 0 #sum of grades
  
  m = 0 #average
  
  averageString = "" # the value that appears in (Nota para Aprovação Final)
  
  for i in student:
    
    naf = 0 #final grade
    
    if(i[0].isnumeric()):
      if(int(i[2]) <= numberClass):
        
        grade = (int(i[3]) + int(i[4]) + int(i[5]))
        
        m =(((grade)/3)/10) # set average
        
        if(m >= 7):
          averageString = "Aprovado"
          
        elif(m >=5 and m < 7):
          averageString = "Exame Final"
          naf = f"Maior que {round(10-m)}"
          
        else:
          averageString = "Reprovado por Nota"
          
      else:
        averageString = "Reprovado por Falta"
        
      
      print(f"The student's average grade: {i[1]} is {m:.2f} and had {i[2]} absences")
        
      try:
        
        i[6] = averageString
        i[7] = naf
        
      except:
        
        i.append(averageString)
        i.append(naf)
        
  return student





if __name__ == "__main__": #execute the function
  main()
  
