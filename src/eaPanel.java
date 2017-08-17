import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.SwingConstants;

/**
 * Panel class for the GUI of Excel-Automation software.
 * This is where the components of the GUI are programmed.
 * @author Nik Zulhilmi Nik Fuaad
 *
 */
public class eaPanel extends JPanel {
	
	//Text Fields
	private static JTextField TAMsTextField;
	private static JTextField file1TextField;
	private static JTextField file2TextField;
	private static JTextField columnsTextField;
	private static JTextField outputNameTextField;
	
	//Buttons
	private JButton infoButton;
	private JButton extractButton;
	private JButton closeButton;
	
	//Labels
	private JLabel TAMLabel;
	private JLabel fileLabel;
	private JLabel columnsLabel;
	private JLabel outputNameLabel;
	
	//Strings
	private static String TAMsString = "";
	private static String file1String = "C:/Users/t-ninikf/Downloads/ExcelSheets/HanimBook.xlsx";
	private static String file2String = "C:/Users/t-ninikf/Downloads/ExcelSheets/Book2.xlsx";
	private static String columnsString = "";
	private static String outputNameString = "output";
	
	//Dialogs
	private static JDialog infoDialog = new JDialog();
	private static JDialog finishDialog = new JDialog();
	private static JDialog fileErrorDialog = new JDialog();
	private static JDialog inputErrorDialog = new JDialog();
	
	//Others
	private static boolean logFileBoolean = false;
	private static JCheckBox logCheckBox;
	
	private static List<String> columns = new ArrayList<String>(Arrays.asList(
			"TAM", "SSR", "Organization Name", "Contract Title", 
			"Contract Number", "Schedule Name", "Value", "TPID",
			"Service Name", "Start Date", "End Date"));
	
	private static List<String> TAMs = new ArrayList<String>(Arrays.asList(
			"abusmt", "alea", "amazahar", "colee",
			"easonlau", "taufiqo", "tuchong", "mseng", "gurushr",
			"huzaidim", "iansu", "jhew", "kansiva", "kkphoon",
			"nabinti", "paerun", "sivask", "superuma"));
	
	public eaPanel() {
		//GUI for TAMs
		TAMLabel = new JLabel("TAMs:");
		TAMLabel.setFont(new Font("Serif", Font.BOLD, 28));
		TAMLabel.setBounds(100, 20, 300, 30);
		
		TAMsString += TAMs.get(0);
		for(int i = 1; i < TAMs.size(); i++) {
			TAMsString += ", ";
			TAMsString += TAMs.get(i);
		}
		
		TAMsTextField = new JTextField();
		TAMsTextField.setFont(new Font("Arial", Font.PLAIN, 24));
		TAMsTextField.setBounds(50, 60, 700, 50);
		TAMsTextField.setText(TAMsString);
		TAMsTextField.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//set TAMs list
				TAMsString = TAMsTextField.getText();
			}
		});
		
		
		//GUI for Files
		fileLabel = new JLabel("Excel files: ");
		fileLabel.setFont(new Font("Serif", Font.BOLD, 28));
		fileLabel.setBounds(100, 130, 300, 30);
		
		file1TextField = new JTextField();
		file1TextField.setFont(new Font("Arial", Font.PLAIN, 24));
		file1TextField.setBounds(50, 170, 700, 50);
		file1TextField.setText(file1String);
		file1TextField.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//set file 1 path
				file1String = file1TextField.getText();
			}
		});
		
		file2TextField = new JTextField();
		file2TextField.setFont(new Font("Arial", Font.PLAIN, 24));
		file2TextField.setBounds(50, 225, 700, 50);
		file2TextField.setText(file2String);
		file2TextField.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//set file 2 path
				file2String = file2TextField.getText();
			}
		});
		
		
		//GUI for columns
		columnsLabel = new JLabel("Columns: ");
		columnsLabel.setFont(new Font("Serif", Font.BOLD, 28));
		columnsLabel.setBounds(100, 290, 300, 30);
		columnsString += columns.get(0);
		for(int i = 1; i < columns.size(); i++) {
			columnsString += ", ";
			columnsString += columns.get(i);
		}
		
		columnsTextField = new JTextField();
		columnsTextField.setFont(new Font("Arial", Font.PLAIN, 24));
		columnsTextField.setBounds(50, 330, 700, 50);
		columnsTextField.setText(columnsString);
		columnsTextField.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//set file 1 path
				columnsString = columnsTextField.getText();
			}
		});
		
		
		//GUI for output file name
		outputNameLabel = new JLabel("Output file name: ");
		outputNameLabel.setFont(new Font("Serif", Font.BOLD, 28));
		outputNameLabel.setBounds(100, 400, 300, 30);
		
		outputNameTextField = new JTextField();
		outputNameTextField.setFont(new Font("Arial", Font.PLAIN, 24));
		outputNameTextField.setBounds(50, 440, 700, 50);
		outputNameTextField.setText(outputNameString);
		outputNameTextField.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//set file 1 path
				outputNameString = outputNameTextField.getText();
			}
		});
		
		
		//GUI for buttons
		infoButton = new JButton("Info");
		infoButton.setFont(new Font("Arial", Font.PLAIN, 24));
		infoButton.setBounds(50, 550, 150, 50);
		infoButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//opens info about the program
				if(!infoDialog.isShowing()) {
					createDialog();
				}
			}
		});
		
		logCheckBox = new JCheckBox("Log to text file");
		logCheckBox.setSelected(false);
		logCheckBox.setBounds(50, 610, 120, 60);
		
		extractButton = new JButton("Extract");
		extractButton.setFont(new Font("Arial", Font.PLAIN, 24));
		extractButton.setBounds(200, 550, 150, 50);
		extractButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				TAMsString = TAMsTextField.getText();
				columnsString = columnsTextField.getText();
				file1String = file1TextField.getText();
				file2String = file2TextField.getText();
				outputNameString = outputNameTextField.getText();
				
				//check if files are ok
				if(fileChecker(file1String, file2String) && !checkFileOpened(file1String)
						&& !checkFileOpened(file2String) && !checkOutputFile(outputNameString)) {
					
					//Check if check box is selected
					if(logCheckBox.isSelected()) {
						logFileBoolean = true;
					}
					else {
						logFileBoolean = false;
					}
					
					//run excel extraction
					extract();
				}
				else {
					notifyFile();
				}
			}
		});
		
		closeButton = new JButton("Close");
		closeButton.setFont(new Font("Arial", Font.PLAIN, 24));
		closeButton.setBounds(350, 550, 150, 50);
		closeButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//close the program (window, frame, and panel)
				//frame.dispose();
				System.exit(0);
			}
		});
		
		setupPanel();
	}
	
	/**
	 * Adds all the components to the panel.
	 */
	private void setupPanel() {
		this.setLayout(null);
		this.add(TAMLabel);
		this.add(TAMsTextField);
		this.add(fileLabel);
		this.add(file1TextField);
		this.add(file2TextField);
		this.add(columnsLabel);
		this.add(columnsTextField);
		this.add(outputNameLabel);
		this.add(outputNameTextField);
		this.add(infoButton);
		this.add(extractButton);
		this.add(closeButton);
		this.add(logCheckBox);
	}
	
	/**
	 * Begin extraction by calling the excelCompare main method.
	 */
	private static void extract() {
		try {			
			if(stringChecker(TAMsString) && stringChecker(columnsString)) {
				columns = convertColumns(columnsString);
				TAMs = convert(TAMsString);
				
				excelCompare.main(null);
			}
			else {
				//print error message
				notifyInput();
			}
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * To check if the string input is in the correct format.
	 * @param s String to be checked.
	 * @return Returns false if string is in wrong format, true otherwise.
	 */
	public static Boolean stringChecker(String s) {
		boolean b1 = true;
		
		s = s.replaceAll("\\s+", ""); //remove all the white spaces
		
		if(s.charAt(0) == ' ') { //check if first character is a comma
			b1 = false;
		}
		
		for (int i = 0; i < s.length() - 2; i++) { // check if there's two commas side by side
			if (s.charAt(i) == ',' && s.charAt(i + 1) == ',') {
				b1 = false;
			}
		}
		
		String temp = s.replaceAll(",", ""); //remove all the commas
		
		if(!temp.matches("[a-zA-Z]+")) { //check if there is only letters
			b1 = false;
		}
		
		return b1;
	}
	
	/**
	 * Convert the string (input from the TextField) into a list.
	 * The list is to be passed to excelCompare to perform the extraction.
	 * @param s String from the TextField.
	 * @return Returns a list where the string is separated by commas.
	 */
	public static List<String> convert(String s) {
		s = s.replaceAll("\\s+", ""); //remove all the white spaces
		List<String> split = Arrays.asList(s.split(",")); //split string into elements separated by ','
		
		return split;
	}
	
	/**
	 * Same as convert, but for columns. This is because the column names might contain more than one word.
	 * Hence removing white spaces will cause problems. We don't want problems.
	 * @param s String from the TextField.
	 * @return Return a list where the elements are separated by commas.
	 * 			Any remaining white spaces at the front/end of the elements will be removed.
	 */
	public static List<String> convertColumns(String s) {
		List<String> split = Arrays.asList(s.split(","));
		
		for(int i = 0; i < split.size(); i++) { //remove white spaces at the front/end of the string
			if(split.get(i).charAt(0) == ' ') { //check first character of the string
				split.set(i, split.get(i).substring(1));
			}
			if(split.get(i).charAt(split.get(i).length() - 1) == ' ') { //check last character of the string
				split.set(i, split.get(i).substring(0, split.get(i).length() - 2));
			}
		}
		
		return split;
	}
	
	/**
	 * Check if the files exist and not a directory.
	 * Files are to be processed.
	 * @param f1 File to be processed.
	 * @param f2 File to be processed.
	 * @return Returns true if the files exist and not a directory. False otherwise.
	 */
	public static boolean fileChecker(String f1, String f2) {
		boolean b1 = false;
		
		File file1 = new File(f1);
		File file2 = new File(f2);
		
		if(file1.exists() && file2.exists() && !file1.isDirectory() && !file2.isDirectory()) {
			b1 = true;
		}
		
		return b1;
	}
	
	/**
	 * Pops up a window saying there's a problem with the files.
	 * Problem could be files don't exist or they are directories.
	 */
	public static void notifyFile() {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				// TODO Auto-generated method stub
				JLabel fileErrorLabel = new JLabel("<html>Check if:"
						+ "<br>1. Files exist"
						+ "<br>2. Files are not directories/folders"
						+ "<br>3. Files are closed before running the program</html>", 
						SwingConstants.CENTER);
				fileErrorLabel.setFont(new Font("Arial", Font.PLAIN, 20));
				
				fileErrorDialog.add(fileErrorLabel);
				fileErrorDialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
				fileErrorDialog.setTitle("Error: Files");
				fileErrorDialog.setSize(new Dimension(400, 300));
				fileErrorDialog.setVisible(true);
				fileErrorDialog.setLocationRelativeTo(null);
				fileErrorDialog.setResizable(false);
			}
		});
	}
	
	/**
	 * Pops up a window saying the inputs aren't typed in correctly.
	 */
	public static void notifyInput() {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				// TODO Auto-generated method stub	
				JLabel inputErrorLabel = new JLabel("Make sure you typed in the inputs correctly!", 
						SwingConstants.CENTER);
				inputErrorLabel.setFont(new Font("Arial", Font.PLAIN, 20));
				
				inputErrorDialog.add(inputErrorLabel);
				inputErrorDialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
				inputErrorDialog.setTitle("Error: Inputs");
				inputErrorDialog.setSize(new Dimension(500, 300));
				inputErrorDialog.setVisible(true);
				inputErrorDialog.setLocationRelativeTo(null);
				inputErrorDialog.setResizable(false);
			}
		});
	}
	
	/**
	 * Pops up a window telling the program has finished extracting.
	 */
	public static void notifyFinish() {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				// TODO Auto-generated method stub
				JLabel finishLabel = new JLabel("Extraction finished!", SwingConstants.CENTER);
				finishLabel.setFont(new Font("Arial", Font.PLAIN, 20));
				
				finishDialog.add(finishLabel);
				finishDialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
				finishDialog.setTitle("Complete");
				finishDialog.setSize(new Dimension(300, 200));
				finishDialog.setVisible(true);
				finishDialog.setLocationRelativeTo(null);
				finishDialog.setResizable(false);
			}
		});
	}
	
	/**
	 * Pops up a window where it contains all the information about the program.
	 */
	private static void createDialog() {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				// TODO Auto-generated method stub
				JTextPane infoLabel = new JTextPane();
				infoLabel.setContentType("text/html");
				infoLabel.setText("<html><left><font size=\"6\">"
						+ "Things to check before extracting: "
						+ "<br>1. There is only ONE excel sheet/tab in one excel file."
						+ "<br>2. There are only letters/alphabets in TAMs and columns fields."
						+ "<br>3. For file path, use forward slashes '/' instead of backward '\\'."
						+ "<br>		ie. C:/Users/t-ninikf/file.exe"
						+ "<br>4. Make sure column names are correct."
						+ "<br>5. Output file name must be valid characters:"
						+ "<br>		No / \\ : * ? \" | < > symbols."
						+ "<br>6. All files to be read (excel sheet from Unicorn and tracker excel sheet)"
						+ " and output file MUST be closed before running the program."
						+ "<br>"
						+ "<br>Info:"
						+ "<br>-First file path is for the Unicorn excel sheet (sheet to be"
						+ " extracted), second is for the tracker sheet."
						+ "<br>-Make sure to update your Java software often."
						+ "<br>-Output file will be in the same directory where the"
						+ " application is executed."
						+ "<br>-Leave the 'Log to text file' box unchecked. Unless there is a problem and the program "
						+ "needs to be debugged."
						+ "<br>"
						+ "<br>Source code: <a href=\"https://github.com/nzulhilmi/Excel-Automation0.0\">"
						+ "https://github.com/nzulhilmi/Excel-Automation0.0</a>"
						+ "<br>"
						+ "<br>More information about the API used to create this program: "
						+ "<a href=\"https://poi.apache.org/spreadsheet/index.html\">"
						+ "https://poi.apache.org/spreadsheet/index.html</a>"
						+ "<br>"
						+ "<br>Any problems please email nzulhilmi94@gmail.com or call:"
						+ "<br>+6011-39377179 / +44 7843132106 (Whatsapp)"
						+ "<br>"
						+ "<br> One of the methods might not work on Linux machines. Please do read the instructions above."
						+ "<br>"
						+ "<br>Visit the source code website for full documentation with Javadoc comments."
						+ "<br>Recommended PC requirements: 4GB RAM, 2048x1536 screen resolution or better."
						+ "<br>Runs on any OS as long as Java is installed."
						+ "<br>"
						+ "<br>"
						+ "<br>by Nik"
						+ "</font></left></html>");
				//infoLabel.setFont(new Font("Arial", Font.PLAIN, 20));
				infoLabel.setEditable(false);
				infoLabel.setMargin(new Insets(20, 20, 20, 20));
				
				infoDialog.add(infoLabel);
				infoDialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
				infoDialog.setTitle("Info");
				infoDialog.setSize(new Dimension(1100, 1100));
				infoDialog.setVisible(true);
				infoDialog.setResizable(false);
			}
		});
	}
	
	/**
	 * Check if a file is already opened.
	 * NOTE: This method will not work on Linux machines. Works on other platforms.
	 * @param s The path of the file to be checked.
	 * @return Returns true if file is already opened.
	 */
	public static boolean checkFileOpened(String s) {
		boolean b1 = false;
		
		File file = new File(s);
		
		//Try to rename the file with the same name
		File sameName = new File(s);
		
		if(file.renameTo(sameName)) {
			//File is closed. Do nothing
		}
		else {
			//File is opened
			b1 = true;
		}
		
		return b1;
	}
	
	/**
	 * Check if output file exists, and if so, check if it is already opened.
	 * @param s File name to be checked.
	 * @return Returns true if file exists and opened, false otherwise.
	 */
	public static boolean checkOutputFile(String s) {
		boolean b1 = false;
		
		String name = s + ".xlsx";
		
		File file = new File(name);
		
		if(file.exists()) {
			if(checkFileOpened(name)) {
				b1 = true;
			}
		}
		
		return b1;
	}
	
	//GET and SET methods
	public static void setTAMs(List<String> list) {
		TAMs = list;
	}
	
	public static List<String> getTAMs() {
		return TAMs;
	}
	
	public static void setFile1(String s) {
		file1String = s;
	}
	
	public static String getFile1() {
		return file1String;
	}
	
	public static void setFile2(String s) {
		file2String = s;
	}
	
	public static String getFile2() {
		return file2String;
	}
	
	public static void setColumns(List<String> list) {
		columns = list;
	}
	
	public static List<String> getColumns() {
		return columns;
	}
	
	public static void setOutputFileName(String s) {
		outputNameString = s;
	}
	
	public static String getOutputFileName() {
		return outputNameString;
	}
	
	public static boolean getLogFileBoolean() {
		return logFileBoolean;
	}
}