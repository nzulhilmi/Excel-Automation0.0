import javax.swing.JFrame;

public class eaFrame extends JFrame{
	private eaPanel currentPanel;
	
	public static void main(String[] args) {
		eaFrame GUI = new eaFrame();
		GUI.setVisible(true);
		GUI.setSize(1100,700);
		
		GUI.setDefaultCloseOperation(EXIT_ON_CLOSE);
	}
	
	public eaFrame() {
		currentPanel = new eaPanel();
		
		setupFrame();
	}
	
	private void setupFrame() {
		getContentPane().setLayout(null);
		this.setContentPane(currentPanel);
	}
}
