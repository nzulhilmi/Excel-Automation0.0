import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JPanel;
//import javax.swing.JTextArea;
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
	private String file1String = "";
	private String file2String = "";
	private String columnsString = "";
	private String outputNameString = "";
	
	public eaPanel() {
		//GUI for TAMs
		TAMLabel = new JLabel("TAMs:");
		TAMLabel.setBounds(100, 20, 300, 30);
		
		TAMsTextField = new JTextField();
		TAMsTextField.setBounds(50, 60, 400, 50);
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
		
		columnsTextField = new JTextField();
		columnsTextField.setBounds(50, 350, 400, 50);
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
				try
				{
					
				}
				catch(Exception e1)
				{
					e1.printStackTrace();
				}
			}
		});
		
		extractButton = new JButton("Extract");
		extractButton.setBounds(250, 550, 150, 50);
		extractButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//run excel extraction
				try
				{
					
				}
				catch(Exception e1)
				{
					e1.printStackTrace();
				}
			}
		});
		
		closeButton = new JButton("Close");
		closeButton.setBounds(400, 550, 150, 50);
		closeButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent e)
			{
				//close the program (window, frame, and panel)
				try
				{
					
				}
				catch(Exception e1)
				{
					e1.printStackTrace();
				}
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
	
	//Get and set methods
	
}
