Monthly Emailer is a Windows Application using .NET Framework 4.7.2. 
The application when run will pair a list of users (provided by a csv file). The application will output an excel file which states which users have been paired together. 
The code uses all the previous excel from this application to make sure the same users don't get paired together again. 
The code will use the active users outlook to open and send an email to the paired users using a template html file. 

The code is written for the csv file to be in the following format
-No header 
column A: user name
column B: company
column C: title
column D: email