# Deployment
We are not responsible for deploying our project to its final, production environment. When we write code, we commit it to a Git repository (hosted on Azure Devops) that we share with OMES. At some point in the future, they will probably incorporate our solution into production. This arrangement was recommended to us by our OMES contact, who gave us access to the shared repository.

If/when OMES does decide to deploy our code to production, it will be as easy as including our ExcelWrapper library as a reference in the HCI website, replacing each instance of "Excel" with "ExcelWrapper", and then changing a few method calls where our library differs syntactically from Excel.

For example, one operation performed in Excel has the form
```wks.UsedRange.Rows.Count```. In ExcelWrapper, the same operation is performed with ```wks.UsedRangeRowsCount```.

## Update to Design
When the design assigment came due, we had just had our old project cancelled, so we couldn't submit a design. Now that our design is fleshed out, we will summarize it here.

Our task is to change some code so that it no longer relies on Microsoft Excel. Instead, the code must create and read spreadsheets using free, open source alternatives.

Our solution is to introduce a library called ExcelWrapper that exposes all the Excel methods used in the OMES code, but behind the scenes implements those methods with the free software SpreadsheetLight. Then, everywhere an Excel method call is made, it can be replaced with the equivalent method call in ExcelWrapper.
