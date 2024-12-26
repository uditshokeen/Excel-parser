# Excel-parser
This Flask application allows users to upload an Excel workbook containing one or more worksheets. The application processes the uploaded Excel file, handles merged cells, and converts the data from each worksheet into a JSON format. It returns a structured JSON response where each worksheet is represented as a separate key in the JSON object.

Key features:
1.	File upload functionality:
•	The application accepts Excel files through a simple web form.

2.	Support for multiple worksheets:
•	The application reads all worksheets from the Excel workbook using ‘pandas.read_excel’ with ‘sheet_name=None’.

3.	Handling merged cells:
    •	Missing values caused by merged cells in excel are addressed using:
	      -Ffill (Forward Fill): Propagates values downward.
        -Bfill (Backward Fill): Propagates values upwards.

4.	JSON conversion:
    •	Each worksheet is converted into a JSON array of records, where:
        -Keys: Column headers
       	-Values: Corresponding row data

5.	Error handling:
•	The application gracefully handles errors during file upload or processing and provides informative error messages.

