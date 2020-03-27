# Testing
Our automated tests can be found in the `ExcelWrapperTests` project. The ones we have currently written are system tests, though some of them may be granular enough to be loosely described as unit tests. Our tests use the `NUnit` framework. Currently, we have 20 of them (or 27 if different test cases are counted as different tests). 

For our project, we are creating a library that replicates a subset of Microsoft's .NET `Excel` library using open source alternatives. Therefore, each of our tests verifies that our library matches `Excel` in some behavior.

Prior to running each test, we invoke a Setup routine that prepares an `Excel` worksheet, along with a Worksheet object in our own library. Then, for the test itself, we perform the same operation on the two worksheets and Assert that we get the same results. Finally, we invoke a Teardown routine to clean up the resources.

We are pretty far from complete code coverage, but we have chosen our tests to cover what we believe are the most nuanced, complicated situations. All of the tests we have are passing.

In addition to these automated tests, we have created a very simple GUI application for manually testing our library. This application lives in the `ExcelWrapperTester` project.

We are hoping to soon receive a sort of acceptance test from our contact at OMES (and have reached out regarding that). This test would be in the form of a database with sample data, and a spreadsheet file that our application should be able to generate, given the database data.