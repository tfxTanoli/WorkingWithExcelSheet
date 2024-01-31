This project gives you the idea of how to work with excel spreadsheet using python with millions of data.
The main thing is that first you have to create a excel file manually in Microsoft Excel and then save that file into your project.
You must download the openpyxl in your project press ALT+F12 to open the terminal and just write "pip install openpyxl"
Mostly people do not do the 2nd step , they just directly copy and paste the code this will give the errors. So , i recommend you to follow the 2nd step.
When you save the workbook and run the program it creates the file successfully in your project directory but if you again run the program it will give you an error. To remove this error you just comment the line which have workbook.save("filename") or wb.save("file") method.
