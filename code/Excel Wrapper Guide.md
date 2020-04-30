## What is this?
We created the ExcelWrapper library to eliminate the dependency of the HCI website on Excel. Specifically, we wanted to get rid of the references to Excel in two places:

1. ```Website1/App_Code/ErrFunctions.vb``` (in the ```MakeIndivSpreadSheets``` function)
2. ```Website1/Import_xls.aspx.vb``` (in the ```XLS_to_TXT``` function)

We did not fully eliminate the dependency on Excel (see the [Moving Forward](##moving-forward) section), but we got most of the way there, with clear next steps. The code we wrote is in 3 separate folders (really, Visual Studio projects). We will go over each of them individually.

## The ```ExcelWrapper``` Project
The ```ExcelWrapper``` folder contains our primary library code: A set of C# classes. Most of these classes match some class from Microsoft's Excel library, in name and behavior. 

Although externally our library behaves the same way as Excel, internally it relies on SpreadsheetLight (a free, open source tool) to interact with spreadsheets. Therefore, by replacing each instance of ```"Excel"``` in the HCI website code with ```"ExcelWrapper"```, we can remove the dependency on Excel. This approach, of swapping out Excel for an equivalent wrapper, allows us to eliminate Excel without understanding the complicated business logic of ```ErrFunctions.vb``` and ```Import_xls.aspx.vb```.

Well, we aren't being completely honest here. There are a few cases where the syntax of ```ExcelWrapper``` differs from the syntax of ```Excel```, though ultimately the functionality is the same. For example, in ```Excel```, you freeze panes in a Worksheet object ```wks``` by calling: 
```
wks.Application.ActiveWindow.FreezePanes = True
```
In ```ExcelWrapper```, the same thing is achieved by:
```
wks.FreezePanes = True
```
These small syntactical differences are rare. Most of them can be seen by examining the changes we made to ```ErrFunctions.vb```. Specifically, look at the diff for commit ```e5c4f```, where we swap out ```Excel``` for ```ExcelWrapper```.

## The ```ExcelWrapperTests``` Project 
The ```ExcelWrapperTests``` folder contains NUnit tests for some of the ```ExcelWrapper``` functionalities. Each test follows the same pattern:

1. Instantiate an ```Excel``` application and an ```ExcelWrapper``` application.
2. Perform the same actions on the two applications.
3. Verify that the results are the same.

We test this way because our goal is for things to work exactly the same when we replace ```Excel``` with ```ExcelWrapper```. All the tests are passing for us as of the writing of this guide.

## The ```ExcelWrapperTester``` Project
The ```ExcelWrapperTester``` folder contains a Visual Basic project you can use to manually test out ```ExcelWrapper```. You can make a spreadsheet with ```ExcelWrapper``` code, then run the GUI to specify a path where the spreadsheet will be saved as .xlsx. We found this useful during devlopement.

## Moving Forward
The only other code we wrote, besides the stuff in the 3 folders, was the update to ```ErrFunctions.vb``` to swap out ```Excel``` with ```ExcelWrapper```. OMES wasn't able to provide us with a testing environment to actually try out the new  ```ErrFunctions.vb```, so it's likely the code still has a few bugs. Although, as far as we know, it could perfectly generate the error spreadsheets as-is. 

There are, however, exactly two behaviors in ```ErrFunctions.vb``` that we were not able to implement in our wrapper library, due to limitations of SpreadsheetLight. These were password protecting the worksheet and saving to TextMSDOS format. When either of these is attemped, we log an error message to console but otherwise do nothing. Handling these functionalities may require using another library besides SpreadsheetLight. That is, you could save the spreadsheet with SpreadsheetLight, open it with a more powerful library, and then perform the desired action.

 (Speaking of more powerful libraries, we were limited to using free options, but we did find some great paid alternatives. Gembox and SpreadsheetGear looked very promising, with more options than SpreadsheetLight and syntax already close to Excel's. It may be worth purchasing one of the paid options for a simpler, cleaner solution.)

There are still several functionalities in ```Import_xls.aspx.vb``` that our wrapper doesn't handle, so we didn't swap out ```Excel``` for ```ExcelWrapper``` in that file. However, we do want to mention one difficult line of code from that file, and how our wrapper can replace it. Specifically, that file uses the command:
```
wks.UsedRange.Rows.Count
```
To avoid having to implement UsedRange, our library instead provides the command
```
wks.UsedRangeRowsCount
```
which manages to do the same thing by a somewhat obscure feature of SpreadsheetLight.

