import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A program which extracts data from Unicorn excel file and processes it alongside with one tracker file.
 * Produces a separate output file.
 * @author Nik Zulhilmi Nik Fuaad
 *
 */
public class excelCompare {
	private static List<String> columnsNeeded = new ArrayList<String>();
	
	private static String[] miaColumns = {"TAM", "SISR", "Customer", "Contract Title",
										"Contract", "Schedule Name", "RM Value", "TPID",
										"Service", "Start Date", "End Date"};
	
	private static List<String> TAMs = new ArrayList<String>();
	
	private static String filepath1;
	private static String filepath2;
	private static String outputFile;
	private static boolean logFileBoolean;
	
	//The number represent the indexes (starts at 0)
	final static int[] numberColumns = {6, 7};
			
	public static void main(String[] args) throws IOException {
		if(logFileBoolean) {
			PrintStream out = new PrintStream(new FileOutputStream("log.txt"));
			System.setOut(out);
		}
		
		assign();
		
		File file1 = new File(filepath1);
		File file2 = new File(filepath2);
		long file1Size = file1.length();
		long file2Size = file2.length();
		
		//Check which file size is bigger. Bigger file size is Unicorn.
		if(file1Size > file2Size) {
			execute(file1, file2, filepath2);
		}
		else {
			execute(file2, file1, filepath2);
		}
		
		eaPanel.notifyFinish();
	}
	
	/**
	 * Reads the content of an excel file and populates the data into an ArrayList.
	 * @param file File to be read.
	 * @param excelSheet ArrayList to be populated.
	 * @param sheetName Name of the sheet name in the excel sheet.
	 */
	public static void read(File file, ArrayList<ArrayList<String>> excelSheet, String sheetName) {
		try {
	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	        
	        workbook.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);
	        DataFormatter fmt = new DataFormatter();
	        
	        XSSFSheet sheet;
	        if(sheetName.equals("null")) {
	        	sheet = workbook.getSheetAt(0);
	        }
	        else {
	        	sheet = workbook.getSheet(sheetName);
	        }
        	
        	for(int rn = sheet.getFirstRowNum(); rn <= sheet.getLastRowNum(); rn++) {
        		//Create 2nd-dimension array list
	            ArrayList<String> columns = new ArrayList<String>();
	            
        		Row row = sheet.getRow(rn);
        		if(row == null) {
        			//There is no data in this row. Need to handle appropriately
        			columns.add("null");
        		}
        		else {
        			for(int cn = 0; cn < row.getLastCellNum(); cn++) {
        				Cell cell = row.getCell(cn);
        				if(cell == null) {
        					//This cell is empty/blank/unused, handle appropriately
        					columns.add("null");
        				}
        				else {
        					String cellStr = fmt.formatCellValue(cell);
        					//Do something with the value
        					columns.add(cellStr);
        				}
        			}
        		}
        		excelSheet.add(columns);
        	}
        	//file.close();
        	workbook.close();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Writes an excel file from an ArrayList.
	 * @param file Not actually needed.
	 * @param excelSheet ArrayList which contains all the data to be written.
	 * @param sheetName The name of the sheet for the file.
	 * @param filename Name of the file to be written.
	 * @param filepath2 Not actually needed.
	 */
	public static void write(File file, ArrayList<ArrayList<String>> excelSheet, String sheetName, 
			String filename, String filepath2) {
		
        try {
        	//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Tracker");
			
			int rowNum = excelSheet.size();
			int colNum = excelSheet.get(0).size();
			int rowCount = 0;
			
			for(int i = 0; i < rowNum; i++) {
				XSSFRow row = sheet.createRow(rowCount++);
				
				int colCount = 0;
				
				for(int j = 0; j < colNum; j++) {
					XSSFCell cell = row.createCell(colCount++);
				
					if((j == 6 || j == 7) && i > 0) {
						cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
						int value = Integer.parseInt(excelSheet.get(i).get(j));
						cell.setCellValue(value);
					}
					else {
						cell.setCellValue((String) excelSheet.get(i).get(j));
					}
				}
			}
			
			FileOutputStream outputStream = new FileOutputStream(filename);
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Execute the whole process of extraction.
	 * @param file1 File supposed to be the data from Unicorn.
	 * @param file2 File supposed to be the tracker sheet.
	 * @param filepath2 File where the program will produce the output file. Not needed at the moment.
	 */
	public static void execute(File file1, File file2, String filepath2) {
		ArrayList<ArrayList<String>> excelSheet1 = new ArrayList<ArrayList<String>>();
		ArrayList<ArrayList<String>> excelSheet2 = new ArrayList<ArrayList<String>>();
		
		String sheetName1 = "Detailed Schedule";
		
		read(file1, excelSheet1, "Sheet1");
		
		//check excel sheet
		ArrayList<ArrayList<String>> output = filter(excelSheet1);
		
		int rowSize = output.size();
		
		sortColumns(output);
		renameColumns(output);
		
		for(int i = 1; i < rowSize; i++) {
			//formatDate_(output.get(i), columnSize - 1);
			correctPrice(output.get(i));
			convertLowerCase(output.get(i));
		}
		
		read(file2, excelSheet2, sheetName1);
		ArrayList<ArrayList<String>> output2 = filter2(excelSheet2);
		
		addRows(output, output2);
		removeComma(output2);
		//output2 is the file that contains all of the data

		String outputFileName = outputFile + ".xlsx";
		write(file2, output2, sheetName1, outputFileName, filepath2);
	}
	
	/**
	 * Traverse through the columns of array, get the indexes of the columns needed and store
	 * them in an array.
	 * @param array Array which needs to be checked.
	 * @return Returns and ArrayList of all the indexes.
	 */
	public static ArrayList<Integer> checkColumns(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> columnNumbers = new ArrayList<Integer>();
		
		int columnSize = array.get(0).size();
		
		for(int i = 0; i < columnSize; i++) {
			if(columnsNeeded.contains(array.get(0).get(i))) {
				columnNumbers.add(i);
			}
		}
		
		return columnNumbers;
	}
	
	/**
	 * Traverse through the rows of array, get the indexes of the rows needed and
	 * store them in an ArrayList.
	 * @param array Array which needs to be checked.
	 * @return Returns an ArrayList with the indexes.
	 */
	public static ArrayList<Integer> checkRows(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> rowNumbers = new ArrayList<Integer>();
		
		int rowSize = array.size();
		
		for(int i = 0; i < rowSize; i++) {
			if(TAMs.contains(array.get(i).get(0))) {
				rowNumbers.add(i);
			}
		}
		
		return rowNumbers;
	}
	
	/**
	 * Filter the input array with only the required columns and rows.
	 * Produces a new ArrayList.
	 * @param array Array which needs to be filtered.
	 * @return Returns a filtered ArrayList.
	 */
	public static ArrayList<ArrayList<String>> filter(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> columnNumbers = checkColumns(array);
		ArrayList<Integer> rowNumbers = checkRows(array);
		ArrayList<ArrayList<String>> output = new ArrayList<ArrayList<String>>();
		
		int numberOfRows = array.size();
		int numberOfColumns = array.get(0).size();
		
		for(int i = 0; i < numberOfRows; i++) {
			//Check for row numbers
			if(rowNumbers.contains(i) || i == 0) {
				ArrayList<String> tempRow = new ArrayList<String>();
			
				for(int j = 0; j < numberOfColumns; j++) {
					if(columnNumbers.contains(j)) {
						tempRow.add(array.get(i).get(j));
					}
				}
				
				output.add(tempRow);
			}
		}
		
		return output;
	}
	
	/**
	 * Acts as a helper method for the filtering process.
	 * @param array Array which needs to be filtered.
	 * @return Returns an ArrayList to be used in filter method.
	 */
	public static ArrayList<ArrayList<String>> filter2(ArrayList<ArrayList<String>> array) {
		ArrayList<ArrayList<String>> output = new ArrayList<ArrayList<String>>();
		
		int rowNumbers = array.size();
		
		for(int i = 0; i < rowNumbers; i++) {
			ArrayList<String> temp = new ArrayList<String>();
			for(int j = 0; j < 11; j++) {
				temp.add(array.get(i).get(j));
			}
			output.add(temp);
		}
		
		return output;
	}
	
	/**
	 * Sort the columns according to the second excel sheet (tracker sheet)
	 * @param array Array in which the columns needs to be sorted.
	 */
	public static void sortColumns(ArrayList<ArrayList<String>> array) {
		int columnSize = array.get(0).size();
		int rowSize = array.size();
		
		for(int i = 0; i < columnSize - 1; i++) {	
			if(!(array.get(0).get(i).equals(columnsNeeded.get(i)))) {	
				//look for index number of searched column/string
				for(int j = i + 1; j < columnSize; j++) {
					if(array.get(0).get(j).equals(columnsNeeded.get(i))) {
						for(int m = 0; m < rowSize; m ++) {
							Collections.swap(array.get(m), i, j);
						}
					}
				}
			}
		}
	}
	
	/**
	 * Acts as a helper method for the method formatDate.
	 * @param date Date to be formatted.
	 * @return Returns a date in the form of a string.
	 */
	public static String formatDate(String date) {
		int dateLength = date.length();
		String temp = date.substring(0, dateLength - 2) + 
				"20" + date.substring(dateLength - 2, dateLength);
		
		return temp;
	}
	
	/**
	 * Change the format of the date on excel sheet.
	 * @param array Array in which the dates need to be formatted.
	 * @param columnSize
	 */
	public static void formatDate_(ArrayList<String> array, int columnSize) {
		//int columnSize = array.get(0).size() - 1;
		
		String temp1 = formatDate(array.get(columnSize));
		String temp2 = formatDate(array.get(columnSize - 1));
		array.set(columnSize, temp1);
		array.set(columnSize - 1, temp2);
	}
	
	/**
	 * Rename the columns according to Mia's tracker sheet.
	 * @param array 
	 */
	public static void renameColumns(ArrayList<ArrayList<String>> array) {
		int columnSize = array.get(0).size();
		
		for(int i = 1; i < columnSize; i++) {
			if(!array.get(0).get(i).equals(miaColumns[i])) {
				array.get(0).set(i, miaColumns[i]);
			}
		}
	}
	
	/**
	 * Convert all the aliases/names in column 1 (SISRs) to lowercase.
	 * @param array
	 */
	public static void convertLowerCase(ArrayList<String> array) {
		String temp = array.get(1);
		array.set(1, temp.toLowerCase());
	}
	
	/**
	 * Remove cent digits from the price value.
	 * ie. RM400.00 to RM400
	 * @param array
	 */
	public static void correctPrice(ArrayList<String> array) {
		String temp = array.get(6);
		int length = temp.length();
		array.set(6, temp.substring(0, length - 3));
	}
	
	/**
	 * Compare rows between two ArrayLists.
	 * Returns true if they are equal, false otherwise.
	 * @param array1
	 * @param array2
	 * @return Returns a boolean if they are equal.
	 */
	public static boolean compareRows(ArrayList<String> array1, ArrayList<String> array2) {
		int columnSize = array2.size();
		
		boolean b1 = true;
		for(int i = 0; i < columnSize; i++) {
			if(!array1.get(i).equals(array2.get(i))) {
				b1 = false;
			}
		}
		
		return b1;
	}
	
	/**
	 * Check for a row array2, if the row is already in array1, returns true. False otherwise.
	 * @param array1 Original array.
	 * @param array2 Array which represents a row in the excel sheet.
	 * @return Returns a boolean value.
	 */
	public static boolean checkExistence(ArrayList<ArrayList<String>> array1, ArrayList<String> array2) {
		int rowSize = array1.size();
		
		boolean b1 = false;
		
		for(int i = 0; i < rowSize; i++) {
			if(compareRows(array1.get(i), array2)) {
				b1 = true;
			}
		}
		
		return b1;
	}
	
	/**
	 * Check array2 rows with array1. If there are any missing rows in array2, will add them.
	 * @param array1 Original array (Which contains the original data).
	 * @param array2 New array where new rows will be added on top of the existing rows.
	 */
	public static void addRows(ArrayList<ArrayList<String>> array1, ArrayList<ArrayList<String>> array2) {
		int rowSize = array1.size();
		
		for(int i = 0; i < rowSize; i++) {
			if(!checkExistence(array2, array1.get(i))) {
				array2.add(array1.get(i));
			}
		}
	}
	
	/**
	 * Remove commas from price value.
	 * ie. 20,000 to 20000.
	 * @param array
	 */
	public static void removeComma(ArrayList<ArrayList<String>> array) {
		int rowNum = array.size();
		
		for(int i = 0; i < rowNum; i++) {
			String value = array.get(i).get(6);
			value = value.replace(",", "");
		
			array.get(i).set(6, value);
		}
	}
	
	/**
	 * Create a text file with error messages.
	 * Produces a new text file or overwrite existing ones.
	 * @param errorMessage Message to be written to the text file.
	 */
	public void printError(String errorMessage) {
		
	}
	
	/**
	 * Print ArrayList to system log output. For debugging purposes.
	 * @param array Array to be printed.
	 */
	public static void printArray(ArrayList<ArrayList<String>> array) {
		int rows = array.size();
		int columns = array.get(0).size();
		
		for(int i = 0; i < rows; i++) {
			for(int j = 0; j < columns; j++) {
				System.out.print(array.get(i).get(j) + " | ");
			}
			System.out.println();
		}
	}
	
	/**
	 * Get all the values from eaPanel before executing the extraction.
	 */
	public static void assign() {
		logFileBoolean = eaPanel.getLogFileBoolean();
		columnsNeeded = eaPanel.getColumns();
		TAMs = eaPanel.getTAMs();
		filepath1 = eaPanel.getFile1();
		filepath2 = eaPanel.getFile2();
		outputFile = eaPanel.getOutputFileName();
	}
	
	//GET and SET methods
	public void setTAMs(List<String> list) {
		TAMs = list;
	}
	
	public static List<String> getTAMs() {
		return TAMs;
	}
	
	public void setFile1(String s) {
		filepath1 = s;
	}
	
	public static String getFile1() {
		return filepath1;
	}
	
	public void setFile2(String s) {
		filepath2 = s;
	}
	
	public static String getFile2() {
		return filepath2;
	}
	
	public List<String> getColumns() {
		return columnsNeeded;
	}
	
	public void setColumns(List<String> list) {
		columnsNeeded = list;
	}
	
	public void setOutputFileName(String s) {
		outputFile = s;
	}
	
	public static String getOutputFileName() {
		return outputFile;
	}
}
