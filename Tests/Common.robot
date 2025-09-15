*** Settings ***
Library    SeleniumLibrary
Library    ../Libs/ExcelLibrairie.py

*** Variables ***
${URL}      https://www.facebook.com
${BROWSER}  Chrome
${data}     C:\\Robot_Tests\\Data\\Common.xlsx
${sheet_name}  ENV


*** Test Cases ***
Open Facebook And Search
    Open Facebook And Fill Email     ${data}

*** Keywords ***
Open Facebook And Fill Email
    [Arguments]    ${data}
    Open Browser    ${URL}    ${BROWSER}
    Maximize Browser Window
    ${row}=    read_first_data_row    ${data}    ${sheet_name}
    Log    ${row}    console=True
    SeleniumLibrary.Input Text    xpath=//input[@id='email']     ${row["LOGIN"]}
    SeleniumLibrary.Input Text     xpath=//input[@id='pass']     ${row["PASSWORD"]}
    Click Element    xpath=//Button[@data-testid='royal-login-button']    
    Sleep    3s
    Capture Page Screenshot
    Close Browser