import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

public class eaPanel extends JPanel {
	
	//Text Fields
	private JTextField TAMsTextField;
	private JTextField file1TextField;
	private JTextField file2TextField;
	private JTextField columnsTextField;
	private JTextField outputNameTextField;
	
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
	private String TAMsString = "";
	private static String file1String = "C:/Users/t-ninikf/Downloads/ExcelSheets/HanimBook.xlsx";
	private static String file2String = "C:/Users/t-ninikf/Downloads/ExcelSheets/Book2.xlsx";
	private String columnsString = "";
	private static String outputNameString = "output";
	
	//Others
	private static List<String> TAMs = new ArrayList<String>(Arrays.asList(
			"TAM", "SSR", "Organization Name", "Contract Title", 
			"Contract Number", "Schedule Name", "Value", "TPID",
			"Service Name", "Start Date", "End Date"));
	
	private static List<String> columns = new ArrayList<String>(Arrays.asList(
			"abusmt", "alea", "amazahar", "colee",
			"easonlau", "taufiqo", "tuchong", "mseng", "gurushr",
			"huzaidim", "iansu", "jhew", "kansiva", "kkphoon",
			"nabinti", "paerun", "sivask", "superuma"));
	
	public eaPanel() {
		//GUI for TAMs
		TAMLabel = new JLabel("TAMs:");
		TAMLabel.setBounds(100, 20, 300, 30);
		
		TAMsString += TAMs.get(0);
		for(int i = 1; i < TAMs.size(); i++) {
			TAMsString += ", ";
			TAMsString += TAMs.get(i);
		}
		
		TAMsTextField = new JTextField();
		TAMsTextField.setBounds(50, 60, 400, 50);
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
		fileLabel.setBounds(100, 130, 300, 30);
		
		file1TextField = new JTextField();
		file1TextField.setBounds(50, 170, 400, 50);
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
		file2TextField.setBounds(50, 240, 400, 50);
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
		columnsLabel.setBounds(100, 280, 300, 30);
		
		columnsString += columns.get(0);
		for(int i = 1; i < columns.size(); i++) {
			columnsString += ", ";
			columnsString += columns.get(i);
		}
		
		columnsTextField = new JTextField();
		columnsTextField.setBounds(50, 350, 400, 50);
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
		outputNameLabel.setBounds(100, 390, 300, 30);
		
		outputNameTextField = new JTextField();
		outputNameTextField.setBounds(50, 460, 400, 50);
		outputNameTextField.setText(outputNameString + ".xlsx");
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
		infoButton.setBounds(100, 550, 150, 50);
		infoButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//opens info about the program
				
			}
		});
		
		extractButton = new JButton("Extract");
		extractButton.setBounds(250, 550, 150, 50);
		extractButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//run excel extraction
				extract();
			}
		});
		
		closeButton = new JButton("Close");
		closeButton.setBounds(400, 550, 150, 50);
		closeButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//close the program (window, frame, and panel)
				
			}
		});
		
		
		setupPanel();
	}
	
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
	}
	
	private static void extract() {
		try {
			excelCompare.main(null);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//Get and set methods
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
}
