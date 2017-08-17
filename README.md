# Excel-Automation0.0
Excel Automation program which extracts information from Unicorn Excel File. The program reads in two files, excel file from Unicorn and excel tracker file. The program will produce an output file with the content from excel tracker file and updated data from Unicorn.

## Things to check before running the program: ##
  1. Make sure there is only ONE excel sheet/tab in once excel file.
  2. Make sure there are only letters/alphabets in TAMs and columns fields.
  3. For file path, use forward slashes '/' instead of backward slashes '\'.
     ie. C:/Users/t-ninikf/Downloads/file.exe
  4. Make sure column names are correct.
  5. Output file name must be valid characters. NO / \ : * ? " | < > symbols.
  6. All files to be read (excel sheet from Unicorn, tracker excel sheet) and output file MUST be closed before running the program.



## Info: ##
* First file path is for the Unicorn excel sheet (sheet to be extracted), second is for the tracker sheet.
* Make sure to update your Java software often.
* Output file will be in the same directory/folder where the application is executed.
* One of the methods might not work on Linux machines. Please do read the instruction above.
* Runs on any OS as long as the correct Java version is installed.



## Recommended PC requirements: ##
  1. 4GB RAM
  2. 2048x1536 screen resolution or better.
  3. Java is installed.



## Information for nerds: ##

The following jar file are used:
  1. commons-codec-1.10.jar
  2. commons-collections.4-4.1.jar
  3. commons-logging-1.2.jar
  4. curvesapi-1.04.jar
  5. junit-4.12.jar
  6. log4j-1.2.17.jar
  7. poi-3.16.jar
  8. poi-examples.3.16.jar
  9. poi-excelant-3.16.jar
  10. poi-ooxml-3.16.jar
  11. poi-ooxml-schemas-3.16.jar
  12. poi-scratchpad-3.16.jar
  13. xmlbeans-2.6.0.jar

  In case of not enough heap/memory size because the excel file is too big, just split the excel file into smaller files.
  Or you can type in the following command before running the program:
  
  For 4GB RAM computers:
  ```
  java -Xms1024m -Xmx4096m -jar Excel-Automation.jar
  ```
  For 8GB RAM computers:
  ```
  java -Xms2048m -Xmx8192m -jar Excel-Automation.jar
  ```



## API ##
More information about the API used to create this program: https://poi.apache.org/spreadsheet/index.html



## Contact ##
Any problems please email nzulhilmi94@gmail.com or call +6011-39377179 (Malaysia) / +44 7843132196 (UK & Whatsapp).



by Nik Zulhilmi Nik Fuaad
