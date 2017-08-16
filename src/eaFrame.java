import java.awt.Dimension;

import javax.swing.JFrame;

/**
 * Frame class for the Excel-Automation GUI.
 * @author Nik Zulhilmi Nik Fuaad
 *
 */
public class eaFrame extends JFrame{
	private eaPanel currentPanel;
	
	public static void main(String[] args) {
		eaFrame GUI = new eaFrame();
		GUI.setTitle("Excel Automation");
		GUI.setVisible(true);
		GUI.setSize(900,800);
		GUI.setLocationRelativeTo(null);
		GUI.setMinimumSize(new Dimension(820,700));
		
		GUI.setDefaultCloseOperation(EXIT_ON_CLOSE);
	}
	
	/**
	 * Creates a new panel to setup the frame.
	 */
	public eaFrame() {
		currentPanel = new eaPanel();
		
		setupFrame();
	}
	
	/**
	 * Set the panel on the frame.
	 */
	private void setupFrame() {
		getContentPane().setLayout(null);
		this.setContentPane(currentPanel);
	}
}
