# Excel-Working-with-function
Use the New Name dialog box to create a named range in the Quarter 1 column.
    Select cell B4 and press Ctrl+Shift+Down Arrow to select the entire range in column B.
    Select Formulas→Define Name.
    In the New Name dialog box, in the Name field, verify Quarter_1 is listed.
    From the Scope drop-down menu, ensure that Workbook is selected.
    Ensure that the Refers to field displays the following range reference: =Region!$B$4:$B$7
    Select OK to create the named range.
Use the Name Box to create a named range in the Quarter 2 column.
    Select cells C4:C7 and select the Name Box, and then type Quarter_2 and press Enter.
    Verify the new range name Quarter_2 is listed in the Name Box.
Use the Create from Selection command to create a named range in the Quarter 3 and Quarter 4 columns.
    Select the range D3:E7 and select Formulas→Create from Selection.
    Ensure that the Top row check box is checked and select OK.
    Select the Name Box drop-down arrow and verify that the two additional named ranges exist, confirming that the names appear as expected.
Use the Create from Selection command to create named ranges for the Region rows simultaneously.
    Select the range A4:E7.
    Select Formulas→Create from Selection.
    Ensure that the Left column check box is checked and select OK.
    Verify that Excel created four unique named ranges for the Region rows.
Navigate to a range and verify the correct total.
    From the Name Box drop-down list, select East.
    Note: You may also use the Go To dialog box to navigate to ranges by pressing F5.
    Verify that Excel selected the quarterly values for the range East in B6:E6.
    With this range selected, note the Total for the range in cell F6 and verify that the same total appears on the Status Bar for the Sum function.
Edit the range names for the quarterly columns to make them a bit shorter.
    Select Formulas→Name Manager.
    Select the Quarter_1 named range and select Edit.
    In the Name field, type Qtr_1 and select OK.
    Change the named range Quarter_2 to Qtr_2
    Edit the Quarter_3 and Quarter_4 named ranges to Qtr_3 and Qtr_4, respectively.
    Close the Name Manager dialog box.
    Examine the Name Box and verify that the names have changed as expected.
Check the AutoSave status of the changed document.
    On the title bar, at the top-left corner of the screen, verify that the status of AutoSave is "On" and the Save button indicates refresh.
    The file is stored in OneDrive and is autosaved when any changes are made, so it is not necessary to save the file by selecting the Save button. However, if            your file is stored on your PC, then you need to select the Save button on the title bar to save the changes to it.
Use an existing range in a function.
        Verify that you are on the Region worksheet, select cell F4, and type =SUM(
        Select Formulas→Use in Formula→North.
        Type ) and press Enter to complete the function.
        Verify that the total for the North region is entered.
Enter a range name with the Formula AutoComplete method.
        Select cell F5, if necessary.
        Type =SUM(sou
        From the Formula AutoComplete pop-up menu, double-click South or press Tab.
        Type ) and press Enter to complete the function.
Replace cell references with range names.
        Select the range F6:F7.
        Select the Formulas tab.
        Select the Define Name drop-down arrow and then select Apply Names.
        Deselect Qtr_4.

Note: The Apply Names dialog box has a built-in sticky function. This means that more than one range name may be selected. Simply deselect any range name not needed, or select all ranges and allow Excel to choose the correct range name.
        Select East and West and select OK.
        Verify that the range names East and West are applied to the formulas in cells F6 and F7, respectively.

Determine which function will insert the current date.
        Select the Employees worksheet and verify that cell B3 is selected.
        On the Help tab, select Help.
        Note: In your own environment, you could alternatively press F1, but in the lab environment, F1 shows the help for the browser you are running the lab in.
        In the Help pane, in the Search box, type insert current date and press Enter.
        Select the Help topic Insert the current date and time in a cell.
        Read the Help topic on inserting the current date and close the Help task pane.
        In cell B3, type =TODAY() and press Enter.
Calculate the years of service value for each employee.
        Select cell C10.
        Enter the formula =($B$3-B10)/365 and press Enter.
        Select cell C10 and double-click the AutoFill handle to fill in the remaining years of service through cell C39.
        Note: Remember, the AutoFill handle is the black square in the bottom-right corner of any cell or range, and when you place your mouse pointer on it, it turns into a black plus sign.
Determine the number of employees with over 20 years of service who will receive the award.
        Select cell B5.
        Select Formulas→Insert Function.
        From the Or select a category drop-down list, select Statistical.
        From the Select a function list box, select COUNTIF and select OK.
        Verify that you are selecting the COUNTIF function and not the COUNTIFS function.
        In the Function Arguments dialog box, verify that your cursor is in the Range text box. Select the range C10:C39 and then press Tab.
        In the Criteria text box, type >=20 and select OK.
        Note: Note that Excel has enclosed your criteria in double quotes.
        Verify that 12 employees have a tenure over 20 years.
        Note: Because the current date changes, there may be more than 12 employees with a tenure over 20 years on the date you are completing the course activity.
Determine the number of employees with less than five years of service who need to attend the safety training.
        Select cell B7.
        On the Formula Bar, select Insert Function.
        From the Or select a category drop-down list, select Most Recently Used.
        From the Select a function list box, select COUNTIF and select OK.
        In the Range text box, select the range C10:C39 and press Tab.
        In the Criteria text box, type "<"&B6
        The ampersand (&) character used here is combining the less than (<) operator enclosed in quotes and the value of cell B6 together, for the criteria "<5".
        Select OK to insert the function.
        Verify that two employees have a tenure less than five years.
Enter a function to calculate the 1% goal bonus for employees.
        Select the Bonus worksheet.
        Verify that cell J8 is selected and type =IF(
        On the Formula Bar, select Insert Function.
        In the Logical_test text box, type G8>H8 and press Tab.
        In the Value_if_true text box, type G8*$C$4 and press Tab.
        In the Value_if_false text box, type 0 and select OK.
        AutoFill the formula in cells J9:J11 to calculate the goal bonus for the remaining employees.
        Verify that a goal bonus has been earned by all but one employee.
Enter a formula to calculate the category bonus, 1% of the sales for each category above $85,000, for the employees.
        Select cell K8 and type =$C$4*SUMIF(
        On the Formula Bar, select Insert Function.
        In the Function Arguments dialog box, in the Range text box, type C8:F8 and press Tab.
        In the Criteria text box, type >85,000 and select OK.
        AutoFill the formula in cells K9:K11 to calculate the category bonus for the remaining employees.
        Verify that all employees except one received a category bonus.
Enter a function to calculate the number of times each employee received a category bonus.
        In cell L8, type =COUNTIF(C8:F8,">"&$C$5) and press Enter.
        Note: The ampersand (&) character used here concatenates, or combines, the greater than (>) operator enclosed in quotes and the value of the cell C5 together, joining the criteria argument for Excel to evaluate as >85,000.
        AutoFill the formula in cells L9:L11 to calculate the number of category bonuses for the remaining employees.
        Verify the counts of each category bonus.
Begin a nested formula to test whether employees receive the Winner's Circle vacation.
        Verify that the Bonus worksheet is selected and select cell N8.
        Type =IF(AND( and then, on the Formula Bar, select Insert Function.
        In the Function Arguments dialog box, in the AND function, verify that your cursor is in the Logical1 text box.
        Type J8>0 and press Tab.
        In the Logical2 text box, type L8>1

Add the arguments for the IF portion of the nested function.
        On the Formula Bar, select the IF function.
        Note: The Function Arguments dialog box will change from the AND function arguments to the IF function arguments.
        In the Function Arguments dialog box, in the IF function, select the Value_if_true text box, type "Winner's Circle" and press Tab.
        In the Value_if_false text box, type "" and select OK.
        Note: There are no spaces between the double quotes.
        AutoFill the formula in cells M9:M11 to calculate the honor for the remaining employees.
        Verify that only Mullins is awarded the Winner's Circle achievement.
        
Enter the NETWORKDAYS function to calculate the total work days for the project.
        Select the Project Details worksheet and verify that the project dates match the following:
        Ensure that cell B9 is selected.
        On the Formula Bar, select Insert function.
        In the Insert Function dialog box, from the Or select a category drop-down list, select the Date & Time category.
        From the Select a function list box, select the NETWORKDAYS function and select OK.
        In the Start_date text box, type or select cell B4 and press Tab.
        In the End_date text box, select cell B5 and press Tab.
        In the Holidays text box, select the range B6:B8 and select OK.
        Verify the total work days for the project. 
        
Select the Campus Information worksheet.
Extract the campus code, the first two characters, from the combined field.
        Ensure that cell G2 is selected.
        Select Formulas→Text→LEFT.
        In the Text text box, type F2 and press Tab.
        In the Num_chars text box, type 2 and select OK.
        Verify that the campus code was extracted.
Extract the building code, the third and fourth characters, from the combined field.
        Select cell H2.
        Select Formulas→Text→MID.
        In the Text text box, type F2 and press Tab.
        In the Start_num text box, type 3 and press Tab.
        In the Num_chars text box, type 2 and select OK.
        Verify that the building code was extracted.
Extract the floor code, the last four characters, from the combined field.
        Select cell I2.
        Select Formulas→Text→RIGHT.
        In the Text text box, type F2 and press Tab.
        In the Num_chars text box, type 4 and select OK.
         Verify that the floor code was extracted.
Concatenate the first name and last name in a single field.
        Select cell J2.
        Type =CONC and press Tab to use Formula AutoComplete.
        In the text1 argument, select or type B2 and type a comma ( , )
        In the [text2] argument, type " " and type a comma ( , )
        Note: There is a space between the two quotation marks.
        In the [text3] argument, select or type D2 and type a right parenthesis ) to complete the function and press Ctrl+Enter.
        Verify that the employee name appears in the full name format.
Join the salutation, first name, middle initial, last name, and suffix into a legal name in a single cell.
        Select cell K2.
        Select Formulas→Text→TEXTJOIN.
        In the Delimiter text box, press the Space bar and then press Tab.
        In the Ignore_empty text box, type true and press Tab.
        In the Text1 text box, type A2:E2 and select OK.
AutoFill in the remaining rows of data.
        Select cells G2:K2 and double-click the AutoFill handle of cell K2.
        Verify that the campus, building, floor, full names, and legal names are listed for all personnel.
        Note how the legal name includes a salutation or suffix when they are present (K3 and K7) but doesn't leave awkward spaces when they are empty.

Close the workbook.
