# Excel-Email-automation
Excel &amp; Email automation using Python, it extracts values from excel file and then puts it in an email saving it as draft requiring just a click-of-a-button.

Automation-Using-Python

Author: Neeraj Mishra

Required libraries:
```python
email, openpyxl
```

Install the libraries using pip and command prompt:

```python
pip install email
pip install openpyxl
```

ReadMe for Project: B&QR
For the project in "B&QR", there is a line of code as follows:

```python
python workbook = excel.load_workbook("TestFile.xlsx")
```

So, this TestFile is created by copying the:
____________________
A - ID
B - Name
C - Client
D - Contact
E - Start Date
F - End Date
G - Employee ID
H - Employee Name
I - Bill Check
J - IQR Range
____________________
columns of the workbook.
We just copy these columns to notepad, then from notepad we copy the content and paste to a new workbook.
This is done, as the original file was giving some issues while trying to get access to the columns data.
So, that is all that is required to be done, apart from that, the mail will be generated as a draft email file, where we would just need to add the mail address to whom we want to send the mail to (Demand Contact).
Note: The feature to fill the mail automatically would be done with future update.
This feature has been enabled.

ReadMe for Project: PE&E
This project is more straightforward, you just have to give the file location in the following field:

```python
workbook = excel.load_workbook("TestFile.xlsx")
```

The columns which will be referenced, are:
____________________
A - Unique ID
B - ID
C - PLM ID
D - Employee Name
E - Employee Email
F - End Date
G - Assignment End Date
H - Client
BE - Manager Email
____________________
That is all there is to this project, and everything else is automated.
The "To" field would be filled by the respective manager, body with respective body (although you might wanna check the "Cc" manually).
