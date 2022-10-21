import os
from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.common.action_chains import ActionChains
import numpy as np

# Read Data
os.chdir("D:\selenium")
data = pd.read_excel('new.xlsx')
data = data.replace(np.nan,'',regex=True)

# set working directory 
# os.chdir("D:\Selenium")
url = 'http://192.168.22.12:81/sessions?orgurl=https:%2F%2Fdqfordynamicssandbox.crm11.dynamics.com&currentUserId=CF7BE984-C971-EC11-8941-002248070A49&alwdrurl=que12Btsrt6OGnk2Uxz8r'
# Chromedriver is just like a chrome. you can dowload latest by it website
driver_path = os.path.join(os.getcwd(), 'chromedriver')
s = Service( driver_path)
driver = webdriver.Chrome(service=s) 
driver.get(url)
time.sleep(15)

# Select Element
def selction(element):
    ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).send_keys(Keys.ESCAPE).perform()

# Delete existing Groups
def deleteGroup(clsName):
    rows2 = driver.find_elements(By.CLASS_NAME, str(clsName))
    for row in rows2:
        ActionChains(driver).key_down(Keys.CONTROL).click(row).key_up(Keys.CONTROL).send_keys(Keys.ESCAPE).perform()

# Drag drop element
def dragdrop(draggable, dropable):
    f = open("D:\Selenium\drag_drop.js",  "r")
    javascript = f.read()
    f.close()
    driver.execute_script(javascript, draggable, dropable)
    time.sleep(3)

# Button 
def button(xpath):
    button = driver.find_element(By.XPATH, str(xpath))
    button.click()
    time.sleep(1)

#dataframe for step 1

d =pd.ExcelFile('new.xlsx')
data1 = pd.read_excel(d, 'Step-1')
data1 = data1.replace(np.nan,'')


#Step 1 

# Read data
matchType = data1['Match Type'].unique()
matchType = matchType[0]

entitytype = data1['Entity'].unique()
entitytype = entitytype[0]

sessionType = data1['Session Type'].unique()
sessionType = sessionType[0]



rows = data['Session Name'].unique()

print("rows",rows)
for val in range(len(rows)):

    if rows[val] != '':
        print("session name is",rows[val])
        time.sleep(2)
        path = driver.find_element(By.XPATH, "/html/body/app-root/div/app-session-folders/div/app-session-header/nav/div/div[2]")
        crtSession = path.find_element(By.XPATH, "/html/body/app-root/div/app-session-folders/div/app-session-header/nav/div/div[2]/a")
        crtSession.click()
        search_field = driver.find_element(By.XPATH, "/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[1]/div[1]/div/div[1]/div/div[2]/div/div/div[1]/div/div/input")
        sessionNames=  rows[val].strip()
        tm = time.strftime('%a, %d %b %Y %H:%M:%S ')
        session_name_tm = sessionNames+ tm
        search_field.send_keys(session_name_tm)


        match_type = driver.find_element(By.XPATH,"/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[1]/div[1]/div/div[1]/div/div[2]/div/div/div[2]/div/div/mat-select/div/div[1]").click()
        time.sleep(2)
        # match_type_val = driver.find_element(By.XPATH, "//span[text() =  ' Single Entity']").click()
        match_type_val = driver.find_element(By.XPATH, "//span[text() =' "+matchType+"']").click()


        #Session type

        time.sleep(2)
        session_type = driver.find_element(By.XPATH, "/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[1]/div[1]/div/div[1]/div/div[2]/div/div/div[3]/div[1]/div/mat-select/div/div[2]").click()
        
        time.sleep(1)
        seesionselection = driver.find_element(By.XPATH, "// span[text() =  ' "+sessionType+"']").click()

        #Entity Selection
        time.sleep(2)
        entity_type = driver.find_element(By.XPATH, "/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[1]/div[1]/div/div[1]/div/div[2]/div/div/div[3]/div[2]/div/mat-select/div/div[2]").click()
        
        time.sleep(1)

        

        
        
        etSelection = driver.find_element(By.XPATH, "// span[text() =  ' "+entitytype+"']").click()

        # Next Button selection
        button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[1]/div[2]/div[1]/button[2]")


        time.sleep(20)
        d =pd.ExcelFile('new.xlsx')
        data1_2 = pd.read_excel(d, 'Step-2')
        data1_2 = data1_2.replace(np.nan,'')
        grpforsession = data1_2.loc[data1_2['Session Name']== sessionNames ]['Match Group Name'].unique()
        numOfGrp = grpforsession.tolist()
        if len(grpforsession) != 1:

            # step-2 - Delete Existing group
            time.sleep(15)
            rows1 = driver.find_elements(By.CLASS_NAME, "match-groups")
            for row in rows1:
                ActionChains(driver).key_down(Keys.CONTROL).click(row).key_up(Keys.CONTROL).send_keys(Keys.ESCAPE).perform()

            #Delete button
            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[2]/app-attribute-groups/div[1]/div/div[2]/div/div[1]/h4/small/button[3]")
            time.sleep(3)

            #Ok button
            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-confirm-dialog/div/mat-dialog-actions/button[1]")
                

            # step-2 - Create Group, set flags and Drag and drop   sessionNames
            
            print(numOfGrp,"name of groups")
            for grp in range(len(numOfGrp)):
                if numOfGrp[grp] != '':
                    # add attribute button
                    button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[2]/app-attribute-groups/div[1]/div/div[2]/div/div[1]/h4/small/button[2]")       
                    add = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-add-group-dialog/div/mat-dialog-content/div/div/div/div[2]/mat-form-field/div/div[1]/div/input") 
                    grpName = numOfGrp[grp].strip()
                    add.send_keys(grpName)
                    time.sleep(2)  
                    #Yes button
                    button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-add-group-dialog/div/mat-dialog-actions/button[1]")  

                    
                    # drag and drop
                    temp = data1_2.loc[data1_2['Match Group Name']== numOfGrp[grp] ]['Match Attributes'].reset_index(drop= True)
                    for i in range(len(temp)):
                        drag = temp[i].strip()
                        c =driver.find_elements(By.CLASS_NAME,"child-node")
                        for j in c:
                            if drag == j.text:
                                str(drag)
                                dragElement =driver.find_element(By.XPATH, "// span[contains(text(), '" + drag + "')]") 
                                time.sleep(1)
                    
                        dropElement =driver.find_element(By.XPATH,"// div[contains(text(), '" + numOfGrp[grp] + "')]" )                               
                        dragdrop(dragElement,dropElement)

                        # step2 -  setting flag values and pririty 

            table = driver.find_element(By.CLASS_NAME, "pm_table_wrapper")
            tbody = table.find_element(By.TAG_NAME, "tbody") 
            trow = tbody.find_elements(By.CLASS_NAME, "match-groups")
            for r in trow:
                    innercols = r.find_elements(By.TAG_NAME, "td")
                    rn =innercols[0].text
                    #priority selection
                    grpValStep2=(data1_2.loc[data1_2['Match Group Name'] == rn]) 
                    prior = grpValStep2['Priority'].unique()
                    a = prior[0]
                    if int(a) == 1:
                        setpath = ' Exact '
                    elif int(a) == 2:
                        setpath = ' Very High '
                    elif int(a) == 3:
                        setpath = ' High '
                    elif int(a) == 4:
                        setpath = ' Medium '
                    elif int(a) == 5:
                        setpath = ' Normal '
                    cls = driver.find_elements(By.CLASS_NAME, "match-groups")
                    div = innercols[2].find_element(By.CLASS_NAME, "matSelect")
                    div.click()
                    time.sleep(1)
                    priority = driver.find_element(By.XPATH, "// span[text() =  '" +setpath+ "']")
                    priority.click()

                    field = grpValStep2['Field'].unique()
                    fld = field[0]
                    if fld == 'Cross':
                        toggle = innercols[3]
                        toggle.click()
                    
                    match = grpValStep2['Matching'].unique()
                    mch = match[0]   
                    if mch == 'Matching':
                        toggle = innercols[4]
                        toggle.click()
                
                        
            # next button for  step-2
            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[2]/div/div[1]/button[2]")
        
        else:
            # next button for  step-2
            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[2]/div/div[1]/button[2]")
        


        # Dataframe for step 3
        d =pd.ExcelFile('new.xlsx')
        data2 = pd.read_excel(d, 'Step-3')
        data2 = data2.replace(np.nan,'')


        #3rd stage

        time.sleep(2)
        # Drag Drop for step-3

        grpforsession = data2.loc[data2['Session Name']== sessionNames ]['Group Name'].unique()
        numOfGrp = grpforsession.tolist()
        if len(grpforsession) != 1:
            for val in range(len(numOfGrp)):
                time.sleep(2)
                rule_name = data2['Rule Name'].unique()
                grp1 = data2.loc[data2['Group Name'] == numOfGrp[val] ]['Rule Name']
                td_text = driver.find_element(By.XPATH,"// td[text() =  ' "+numOfGrp[val]+" ']")
                for rule_names in grp1:
                    
                    if rule_names == "Casing Format":
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value = str(parameter1[0])     
                        casing_format = driver.find_element(By.ID, "Case")
                        dragdrop(casing_format,td_text)
                        searchbox = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div/div/div[2]/mat-select")
                        searchbox.click()
                        time.sleep(1)
                        case = driver.find_element(By.XPATH, "// span[contains(text(), ' "+param_value+" ')]")
                        time.sleep(1)
                        case.click()
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                        time.sleep(2)
                    
                    elif rule_names == "Custom Exclude":
                            drag_element= driver.find_element(By.ID, "CustomExclude")
                            dragdrop(drag_element,td_text)
                            time.sleep(1)
                            parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                            param_value1 = str(parameter1[0])
                            left_delimiter = driver.find_element(By.XPATH,"/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[1]/div/div[2]/input")
                            left_delimiter.send_keys(param_value1)
                            parameter2 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter2']).to_list()
                            param_value2 = str(parameter2[0])
                            right_delimiter = driver.find_element(By.XPATH,"/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[2]/div/div[2]/input")
                            right_delimiter.click()
                            right_delimiter.send_keys(param_value2)
                            parameter3 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter3']).to_list()
                            param_value3 = str(parameter3[0])
                            mode = driver.find_element(By.XPATH,"/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[3]/div/div[2]/mat-select/div/div[2]/div")
                            mode.click()
                            select_element = driver.find_element(By.XPATH,"// span[contains(text(), ' "+param_value3+ " ')]" )
                            select_element.click()
                            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                            time.sleep(2)
                        
                    elif rule_names == "Custom Transform":
                        drag_element= driver.find_element(By.ID, "CustomTransform")
                        dragdrop(drag_element,td_text)
                        look_for = driver.find_element(By.XPATH,"/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[1]/div/div[2]/input")
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        look_for.send_keys(param_value1)
                        change_to = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[2]/div/div[2]/input")
                        parameter2 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter2']).to_list()
                        param_value2 = str(parameter2[0])
                        change_to.click()
                        change_to.send_keys(param_value2)
                        parameter3 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter3']).to_list()
                        param_value3 = str(parameter3[0])
                        if param_value3 == "1":
                            check_box = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[3]/div/div[2]/div/div/label/span/span").click() 
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                        time.sleep(2)
                    
                    elif rule_names == "Custom Transform Library":
                        drag_element= driver.find_element(By.ID, "CustomTransformLibrary")
                        dragdrop(drag_element,td_text)
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        cateogary = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div/div/div[2]/mat-select/div/div[2]").click()
                        select_element = driver.find_element(By.XPATH,"// span[contains(text(), ' "+param_value1+" ')]" )
                        select_element.click()
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                        time.sleep(2)

                    elif rule_names == "Extract Letters":
                        drag_element= driver.find_element(By.ID, "ExtractLetters")
                        dragdrop(drag_element,td_text)
                        select_value = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[1]/div/div[2]/mat-select/div/div[2]").click()
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        select_element = driver.find_element(By.XPATH,"// span[text() = '"+param_value1+"']" )
                        select_element.click()
                        no_of_letters = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[2]/div/div[2]/input")
                        no_of_letters.click()
                        parameter2 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter2']).to_list()
                        param_value2 = str(parameter2[0])
                        no_of_letters.send_keys(param_value2)
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                        
                        time.sleep(2)
                    elif rule_names == "Extract Name":
                        drag_element= driver.find_element(By.ID, "ExtractName")
                        dragdrop(drag_element,td_text)
                        select_value = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[1]/div/div[2]/mat-select/div/div[2]").click()
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        select_element = driver.find_element(By.XPATH,"// span[contains(text(), '"+param_value1+"')]" )
                        select_element.click()
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                    
                    elif rule_names == "Extract Word":
                        drag_element= driver.find_element(By.ID, "ExtractWord")
                        dragdrop(drag_element,td_text)
                        select_value = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[1]/div/div[2]/mat-select/div/div[2]").click()
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        select_element = driver.find_element(By.XPATH,"// span[text() = '"+param_value1+"']" )
                        select_element.click()
                        no_of_words = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div/div[2]/div/div[2]/input")
                        no_of_words.click()
                        parameter2 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter2']).to_list()
                        param_value2 = str(parameter2[0])
                        no_of_words.send_keys(param_value2)
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                        time.sleep(2)

                    elif rule_names == "Remove Characters":
                        drag_element= driver.find_element(By.ID, "RemoveChars")
                        dragdrop(drag_element,td_text)
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        time.sleep(0.5)
                        parameter2 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter2']).to_list()
                        param_value2 = str(parameter2[0])
                        time.sleep(0.5)
                        parameter3 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter3']).to_list()
                        param_value3 = str(parameter3[0])
                        parameter4 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter4']).to_list()
                        param_value4 = str(parameter4[0])
                        parameter5 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter5']).to_list()
                        param_value5 = str(parameter5[0])

                        if param_value1 == str(1):
                            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[1]/div/label/span/span")
                        if param_value2 == str(1):
                            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[2]/div/label/span/span")
                        if param_value3 == str(1):
                            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[3]/div/label/span/span")
                        if param_value4 == str(1):
                            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[4]/div/label/span/span")
                        if param_value5 == str(1):
                            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[5]/div/label/span/span")  

                        time.sleep(0.5)
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                    
                    elif rule_names == "Split String":
                        drag_element= driver.find_element(By.ID, "SplitString")
                        dragdrop(drag_element,td_text)
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        time.sleep(0.5)
                        parameter2 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter2']).to_list()
                        param_value2 = str(parameter2[0])
                        time.sleep(0.5)
                        parameter3 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter3']).to_list()
                        param_value3 = str(parameter3[0])
                        deli = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[1]/div/div/div[2]/input")
                        deli.send_keys(param_value1)
                        if param_value2 == '1':
                            checkbox = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[2]/div/div/div[2]/div/div/label/span")
                            checkbox.click()
                        mode = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[3]/div/div/div[2]/mat-select")
                        mode.click()
                        select = driver.find_element(By.XPATH,"// span[text() =  ' "+param_value3+" ']")
                        select.click()
                        time.sleep(0.5)
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")

                    elif rule_names == "Split String":
                        drag_element= driver.find_element(By.ID, "Normalise")
                        dragdrop(drag_element,td_text)
                        parameter1 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter1']).to_list()
                        param_value1 = str(parameter1[0])
                        parameter2 = (data2.loc[data2['Rule Name']== rule_names ]['Parameter2']).to_list()
                        param_value2 = str(parameter2[0])
                        method = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[1]/div/div/div[2]/mat-select")
                        method.click()
                        ele1 = driver.find_element(By.XPATH,"// span[text() =  ' "+param_value1+" ']")
                        ele1.click()
                        time.sleep(1)
                        cat = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-content/div[2]/div/div/div[2]/mat-select")
                        cat.click()         
                        ele = driver.find_element(By.XPATH,"// span[text() =  ' "+param_value2+" ']")
                        ele.click()
                        time.sleep(0.5) 
                        button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-transform-dialog/div/mat-dialog-actions/button[1]")
                        
                    elif rule_names == "Trim string":
                        drag_element= driver.find_element(By.ID, "TrimString")
                        dragdrop(drag_element,td_text)
                    
                    else:
                        print("Invalid Option")

            # Next Button for step-3
            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[3]/div/div[1]/button[2]")

        else:
             # Next Button for step-3
            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[3]/div/div[1]/button[2]")

        # dataframe for step 4
        d =pd.ExcelFile('new.xlsx')
        data3 = pd.read_excel(d, 'Step-4')
        data3 = data3.replace(np.nan,'')


        #4th stage
        time.sleep(2)


        # Set slider in middle
        slider = driver.find_element(By.CLASS_NAME, "mat-slider")
        move = ActionChains(driver)
        move.click_and_hold(slider).move_by_offset(0, 0).release().perform()  
        time.sleep(2)

        threshold = data3.loc[data3['Session Name']== sessionNames ]['threshold'].unique()
    
        val_threshold = threshold[0]
        print("val_threshold",val_threshold)
        actual = 75

        final = int(val_threshold - actual)

        if final > 0:
            for i in range(final):
                move.send_keys(Keys.ARROW_RIGHT).perform()
        else:
            final = -(final)
            for i in range(final):
                move.send_keys(Keys.ARROW_LEFT).perform()


        # Set Flags in Step-4
        val1 = "Include original data when scoring matches"
        val_include_original_data = data3['Include original data when scoring matches'].values
        include_original_data = val_include_original_data[0]
        if include_original_data == 1:
            first_toggle =driver.find_element(By.XPATH, "// label[contains(text(), ' "+val1+" ')]")
            first_toggle.click()
        else:
            pass


        val2 = "Treat two Null fields as 100% match"
        val_treat_two_null = data3['Treat two Null fields as 100% match'].values
        treat_two_null = val_treat_two_null[0]
        if treat_two_null == 1:
            first_toggle =driver.find_element(By.XPATH, "// label[contains(text(), ' "+val2+" ')]")
            first_toggle.click()
        else:
            pass

        val3 ="Value to Null field as 100% match"
        val_value_to_null = data3['Value to Null field as 100% match'].values
        value_to_null = val_value_to_null[0]

        if value_to_null == 1:
            first_toggle =driver.find_element(By.XPATH, "// label[contains(text(), ' "+val3+" ')]")
            first_toggle.click()
        else:
            pass

        # Next Button in Step-4
        button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[4]/div/div[1]/button[2]")



        #5th stage - delete existing groups
        dOne =pd.ExcelFile('new.xlsx')
        data = pd.read_excel(dOne, 'Step-5')
        data = data.replace(np.nan,'')

        grpforsession = data.loc[data['Session Name']== sessionNames ]['DNAF Group Name'].unique()
        dnfgrpname = grpforsession.tolist()
        
        if len(grpforsession) != 1:

            time.sleep(25)
            deleteGroup("dnaf-groups")
            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[5]/app-display-and-autofill-settings/div[1]/div/div[2]/div/div[1]/h4/small/button[3]")
            time.sleep(2)
            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-confirm-dialog/div/mat-dialog-actions/button[1]")


            #dataframe for step 5 - group rules

        

            dgrp = pd.read_excel(dOne, 'Step5GroupRules')
            dgrp = dgrp.replace(np.nan,'')

            #5th stage add new groups
        
            for grp in range(len(dnfgrpname)):

                    if dnfgrpname[grp] != '':
                            # add attribute button
                            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[5]/app-display-and-autofill-settings/div[1]/div/div[2]/div/div[1]/h4/small/button[2]")       
                            add = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-add-group-dialog/div/mat-dialog-content/div/div/div/div[2]/mat-form-field/div/div[1]/div/input") 
                            grpName = str(dnfgrpname[grp]).strip()
                            add.send_keys(grpName)
                            time.sleep(2)  
                            #Yes button
                            button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-add-group-dialog/div/mat-dialog-actions/button[1]") 

                            # ### for group rules ##
            
                            # trow = driver.find_elements(By.CLASS_NAME, "dnaf-groups")
                    
                            # row = trow[grp].find_elements(By.TAG_NAME, "td")[1].click()

                            # time.sleep(2)
                            # grpVal=(dgrp.loc[dgrp['Group Name'] == grpName, 'Group Rules'])
                            # time.sleep(2)

                            # for rules in grpVal:

                                    
                            #         time.sleep(2)      
                            #         grpRls = driver.find_element(By.XPATH, "// div[text() = '"+rules+"']")
                            #         tr = driver.find_element(By.CLASS_NAME, "ng-star-inserted")
                            #         drp = tr.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-rule-dialog/div/div[2]/div[1]/div/div[2]/div/div[2]/div/div/div/table/tbody/tr[2]/td[1]")
                            #         dragdrop(grpRls,drp)
                            #         button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div[1]/div/div[2]/mat-select")
                            #         ele1_value =(dgrp.loc[dgrp['Group Rules'] == rules, 'Parameter 1'])
                            #         ele1_val = (ele1_value.reset_index(drop=True)[0])
                            #         ele1 = driver.find_element(By.XPATH,"// span[text() = ' "+ele1_val+" ']")
                            #         ele1.click() 

                            #         ele2_value =(dgrp.loc[dgrp['Group Rules'] == rules, ' Parameter 2'])
                            #         ele2_val = str(ele2_value.reset_index(drop=True)[0])

                            #         time.sleep(1)
                            #         ele2 = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div[2]/div/div[2]/mat-form-field/div/div[1]/div[3]/input")
                            #         # ele2.click()
                            #         ele2.send_keys(ele2_val)
                            #         #apply btn
                            #         button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-actions/button[1]/span")

                            
                            
                            # # #save button
                            # button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-rule-dialog/div/div[2]/div[2]/div/button[1]")


                            dnaf_attributes = data['DNAF Attributes'].unique()
                            grp1 = data.loc[data['DNAF Group Name'] == dnfgrpname[grp] ]['DNAF Attributes'] 
                            drop = driver.find_element(By.XPATH,"// span[text() =  '"+dnfgrpname[grp]+"']")
                            time.sleep(2)
                            for dnaf_attr in grp1:
                                    time.sleep(2)

                                    drag_element = driver.find_element(By.XPATH,"// span[text() = ' "+dnaf_attr+" ']" )
                                    dragdrop(drag_element,drop)

                            grpVal1=(dgrp.loc[dgrp['Group Name'] == grpName, 'Group Rules'])
                            grpVal = grpVal1.tolist()
                            print("***",grpVal)

                            if len(grpVal) != 0:
                                trow = driver.find_elements(By.CLASS_NAME, "dnaf-groups")
                                    
                                row = trow[grp].find_elements(By.TAG_NAME, "td")[1].click()

                                time.sleep(2)
                                print("grpVal",grpVal)
                                

                                for rules in grpVal:

                                        print("rules",rules)

                                        
                                        time.sleep(2)      
                                        grpRls = driver.find_element(By.XPATH, "// div[text() = '"+rules+"']")
                                        tr = driver.find_element(By.CLASS_NAME, "ng-star-inserted")
                                        drp = tr.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-rule-dialog/div/div[2]/div[1]/div/div[2]/div/div[2]/div/div/div/table/tbody/tr[2]/td[1]")
                                        dragdrop(grpRls,drp)
                                        button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div[1]/div/div[2]/mat-select")
                                        ele1_value =(dgrp.loc[dgrp['Group Rules'] == rules, 'Parameter 1'])
                                        ele1_val = (ele1_value.reset_index(drop=True)[0])
                                        ele1 = driver.find_element(By.XPATH,"// span[text() = ' "+ele1_val+" ']")
                                        ele1.click() 

                                        ele2_value =(dgrp.loc[dgrp['Group Rules'] == rules, ' Parameter 2'])
                                        ele2_val = str(ele2_value.reset_index(drop=True)[0])

                                        time.sleep(1)
                                        ele2 = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div[2]/div/div[2]/mat-form-field/div/div[1]/div[3]/input")
                                        # ele2.click()
                                        ele2.send_keys(ele2_val)
                                        #apply btn
                                        button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-actions/button[1]/span")

                                
                                
                                # #save button
                                button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-rule-dialog/div/div[2]/div[2]/div/button[1]")
                                # time.sleep(5)
                            # button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[5]/div/div[2]/button[2]")
                            # time.sleep(25)



                            else:

                                #Step 5 : Toggle buttons

                    #####       

                                dOne =pd.ExcelFile('new.xlsx')
                                dataAtt = pd.read_excel(dOne, 'Step5AttributeRules')
                                dataAtt = dataAtt.replace(np.nan,'')


                                # to set attribute rules
                                def attrRules(attrName):
                                    rule_val_owner =  dataAtt.loc[dataAtt["DNAF Attributes"] == attrName, "Rules"].to_list()
                                    for ind,i in enumerate(rule_val_owner):  
                                        drag_element =  driver.find_element(By.ID, ''+i+ '')
                                        drop_element =  driver.find_element(By.XPATH,"// td[text() = ' Rules Sequence ']")
                                        dragdrop(drag_element,drop_element)   
                                        rule_val1= dataAtt.loc[dataAtt["DNAF Attributes"] == attrName].reset_index(drop = True)

                                    
                                        try:
                                            try:
                                                #single value (drop-down list)
                                                time.sleep(1)
                                                val1 = driver.find_element(By.XPATH,"/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div[1]/div/div[2]") 
                                                val1.click() 
                                                sp = driver.find_element(By.XPATH, "// span[text() = ' "+rule_val1["Parameter1"][ind]+" ']")
                                                time.sleep(1)
                                                sp.click()
                                            except:           
                                                try:
                                                    try:
                                                    
                                                        time.sleep(1)
                                                        base_score = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[3]/div[2]/span/input")
                                                        base_score.send_keys(rule_val1["Parameter1"][ind])
                                                        # new
                                                        btn = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[4]/div/button").click()
                                                        time.sleep(2)
                                                        driver.find_element(By.XPATH, "/html/body/div[3]/div[6]/div/mat-dialog-container/app-option-set-dialog/div/mat-dialog-content/div[3]/div/mat-table/mat-row/mat-cell[1]/div/mat-checkbox/label/div").click()
                                                        button("/html/body/div[3]/div[6]/div/mat-dialog-container/app-option-set-dialog/div/mat-dialog-actions/button[1]")
                                                        time.sleep(1)
                                                        button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[5]/div/div/table/tbody/tr/td[1]/div/label/span/span")

                                                    except: 
                                                        # for owner type -- hierarchy of value
                                                        base_score = driver.find_element(By.XPATH, " /html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[4]/div[2]/span/input")
                                                        base_score.send_keys(rule_val1["Parameter1"][ind])
                                                        time.sleep(2)
                                                        button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[5]/div/button")
                                                        time.sleep(15)

                                                        button("/html/body/div[3]/div[6]/div/mat-dialog-container/app-lookup-option-set-dialog/div/mat-dialog-content/div[4]/div/mat-table/mat-header-row/mat-header-cell[1]/div/label/span/span")
                                                        time.sleep(2)
                                                        button("/html/body/div[3]/div[6]/div/mat-dialog-container/app-lookup-option-set-dialog/div/mat-dialog-actions/button[1]")

                                                except:
                                                    time.sleep(1)
                                                    val = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[3]/div[2]/div/mat-select/div/div[1]")
                                                    val.click()
                                                    sp = driver.find_element(By.XPATH, "// span[text() = ' "+rule_val1["Parameter1"][ind]+"']")
                                                    sp.click()
                                                
                                                
                                        except:
                                            try:
                                                #single Value (related attributes)
                                                sp = driver.find_element(By.XPATH, "/html/body/div[3]/div[6]/div/div/div/mat-option[3]/span")
                                                time.sleep(1)
                                                sp.click()
                                            except:
                                                # single value (to enter value)
                                                time.sleep(1) 
                                                val =  driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div/div/div[2]/mat-form-field/div/div[1]/div[3]/input")
                                                val.send_keys(rule_val1["Parameter1"][ind])     
                                        time.sleep(2)
                                        if rule_val1["Parameter2"][ind] != '':
                                            try:
                                                val2 = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div[2]/div/div[2]/mat-form-field/div/div[1]/div[3]/input")
                                                val2.send_keys(rule_val1["Parameter2"][ind])
                                                time.sleep(2)
                                            except:

                                                try:
                                                    val2 = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[4]/div[2]/span/input")
                                                    val2.send_keys(rule_val1["Parameter2"][ind])
                                                    time.sleep(2)
                                                except:
                                                    element = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[4]/div[2]/span/input")
                                                    element.send_keys(rule_val1["Parameter2"][ind])
                                            


                                        else:
                                            pass
                                        if rule_val1["Parameter3"][ind]  != "":
                                            try:
                                                val3 = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-content/div/div/div[3]/div/div[2]/mat-form-field/div/div[1]/div[3]/input")
                                                val3.send_keys(rule_val1["Parameter3"][ind])
                                            except:
                                                btn_val = rule_val1["Parameter3"][ind]
                                                #add value
                                                button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-content/div[4]/div[3]/button")
                                                time.sleep(1)
                                            
                                        else:
                                            pass
                                        time.sleep(1)

                                        #apply button
                                        try:
                                            button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-rule-dialog/div/mat-dialog-actions/button[1]")
                                        except:
                                            button("/html/body/div[3]/div[4]/div/mat-dialog-container/app-hierarchical-value-dialog/div/mat-dialog-actions/button[6]")  

                                    # save button
                                    button("/html/body/div[3]/div[2]/div/mat-dialog-container/app-group-rule-dialog/div/div[2]/div[2]/div/button[1]")



                    #####
                                table = driver.find_element(By.CLASS_NAME, "pm_tabe_displayGroupRules")
                                tbody = table.find_element(By.TAG_NAME, "tbody") 
                                trow = tbody.find_elements(By.CLASS_NAME, "pm_collapse_panel") #loop1
                                trow = trow[: len(trow)-1]
                                for r in trow: 
                                        td = r.find_element(By.TAG_NAME, "td")
                                        innerTable = td.find_element(By.TAG_NAME, "table")
                                        innerRows = innerTable.find_elements(By.TAG_NAME, "tr") #loop2
                                        for ir in innerRows:
                                            innercols = ir.find_elements(By.TAG_NAME, "td") #loop3
                                            rn =innercols[0].text
                                            grpVal=(data.loc[data['DNAF Attributes'] == rn]) 
                                            attr_rule_type = innercols[0]
                                            innercols[1].find_element(By.TAG_NAME, "button").click()
                                            time.sleep(2)
                                            img_attr_rule_type = attr_rule_type.find_element(By.TAG_NAME,"img") # type of attribute
                                            attrType = img_attr_rule_type.get_attribute("ng-reflect-message")
                                            time.sleep(2)
                                            rule_val =  dataAtt.loc[dataAtt["DNAF Attributes"] == attr_rule_type.text, "Rules"].to_list()
                                            time.sleep(2)
                                            attrRules(attr_rule_type.text)
                                            
                                            uniq1 = grpVal["Show on DQ4Dyn"].values
                                            uniq1= uniq1[0]  # 1
                                            ip1 = innercols[3].find_element(By.TAG_NAME, "input")
                                            ipVal1 = ip1.get_attribute("ng-reflect-model")   # 0

                                        
                                            if ipVal1 == 'false':
                                                    bval1 = 0
                                            else:
                                                    bval1 = 1
                                            if bval1 != uniq1:
                                                    innercols[3].click()
                                            time.sleep(1)

                                            uniq2 = grpVal["Exclude Update"].values
                                            uniq2= uniq2[0]
                                            
                                            ip2 = innercols[4].find_element(By.TAG_NAME, "input")
                                            ipVal2 = ip2.get_attribute("ng-reflect-model")

                                            if ipVal2 == 'false':
                                                    bval2 = 0
                                            else:
                                                    bval2 = 1
                                            
                                            if bval2 != uniq2:
                                                    innercols[4].click()
                                            time.sleep(1)

                                            uniq3 = grpVal["Use for Auto Exact & Merge"].values
                                            uniq3 = uniq3[0]
                                            ip3 = innercols[5].find_element(By.TAG_NAME, "input")
                                            ipVal3 = ip3.get_attribute("ng-reflect-model")
                                            if ipVal3 == 'false':
                                                    bval3 = 0
                                            else:
                                                    bval3 = 1
                                            if bval3 != uniq3:
                                                    innercols[5].click()

                                            time.sleep(1)
                                            uniq4 = grpVal["Ignore Nulls for Auto Exact & Merge"].values
                                            uniq4 = uniq4[0]
                                            ip4 = innercols[6].find_element(By.TAG_NAME, "input")
                                            ipVal4 = ip4.get_attribute("ng-reflect-model")
                                            if ipVal4 == 'false':
                                                    bval4 = 0
                                            else:
                                                    bval4 = 1
                                            if bval4 != uniq4:
                                                    innercols[6].click()
                                            time.sleep(2)

            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[5]/div/div[2]/button[2]")
            time.sleep(25)
        else:
            button("/html/body/app-root/div/app-create-session/div/div[1]/mat-horizontal-stepper/div[2]/div[5]/div/div[2]/button[2]")
            time.sleep(25)






            
                   

        


        