import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FilenameFilter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.table.DefaultTableModel;

import com.spire.presentation.IAutoShape;
import com.spire.presentation.ISlide;
import com.spire.presentation.ParagraphEx;
import com.spire.presentation.Presentation;

public class FileDialog {

	public static void main(String[] args) throws Exception {
		FileSearcher js =new FileSearcher();
	}

}

class FileSearcher extends JFrame implements ActionListener {
	
	File dir;
	File[] pptxfiles;
	
	JTextField search;
	JFileChooser jfc = new JFileChooser();
	
	JButton btnOpen = new JButton("select folder");
	JButton btnSearch = new JButton("search start");
	JLabel dirpath = new JLabel(" ");
	
	JScrollPane scrolledTable;
	static JTable table;
	
	public FileSearcher() throws Exception {
		this.setSize(800, 500);
		this.setLocation(100, 100);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setLayout(new BorderLayout());
		this.panelinit();
		this.tableinit();
		this.set();
		this.setVisible(true);
	}
	
	public void panelinit() {
		JPanel topPanel1 = new JPanel();
		JPanel topPanel2 = new JPanel();

		topPanel1.add(btnOpen);
		
		JLabel searchlabel = new JLabel("search target : ");
		topPanel2.add(searchlabel);
		search = new JTextField(10);
		topPanel2.add(search);
		topPanel2.add(btnSearch);
		
		JPanel topPanel = new JPanel();
		topPanel.add(topPanel1);
		topPanel.add(dirpath);
		add(topPanel, BorderLayout.WEST);		
		add(topPanel2, BorderLayout.EAST);		
	}
	public void tableinit() {
		
		String header[] = {"file name", "page", "textbox", "textline", "text"};
		DefaultTableModel model = new DefaultTableModel(header,0);
		
		table = new JTable(model);
		table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		scrolledTable = new JScrollPane(table);
		scrolledTable.setBorder(BorderFactory.createEmptyBorder(0,10,0,10));
		
		add(scrolledTable,BorderLayout.SOUTH);
	}
	
	public void set() {
		btnOpen.addActionListener(this);
		btnSearch.addActionListener(this);
		jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		jfc.setMultiSelectionEnabled(false);
	}
	
	//버튼이 눌렸을 경우
	public void actionPerformed(ActionEvent arg0) {
		if (arg0.getSource() == btnOpen) {
			if (jfc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
				dirpath.setText("searching filepath : " + jfc.getSelectedFile().toString());
				dir = jfc.getSelectedFile();
				pptxfiles = dir.listFiles(new FilenameFilter() {
					
					@Override
					public boolean accept(File dir, String name) {
						return name.endsWith("pptx");
					}
				});
			}
		} else if (arg0.getSource() == btnSearch) {
			resetting();
			for (File file : pptxfiles) {
				try {
					GetText(file, search.getText());
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			
		}
	}
	
	// pptx에서 target에 해당하는 text를 추출해서 setting 돌리기
	public static void GetText(File file, String target) throws Exception {
		int pagenum = 0;
		int textnum = 0;
		int textline = 0;
		Presentation ppt = new Presentation();
		ppt.loadFromFile(file.toString());
		for (Object slide : ppt.getSlides()) {
			pagenum++;
			textnum = 0;
			for (Object shape : ((ISlide) slide).getShapes()) {
				textnum++;
				textline = 0;
				if (shape instanceof IAutoShape) {
					for (Object tp : ((IAutoShape) shape).getTextFrame().getParagraphs()) {
						String txt = ((ParagraphEx) tp).getText();
						textline++;
						if (!txt.contains(target)) {continue;}
						setting(file.getName(),String.valueOf(pagenum),String.valueOf(textnum),
								String.valueOf(textline),((ParagraphEx) tp).getText());
					}
				}
			}
		}
	}
	
	// table 초기화
	public static void resetting() {
		
		DefaultTableModel model = (DefaultTableModel)table.getModel();
		int rownum = model.getRowCount();
		for (int i=0; i<rownum;i++) {model.removeRow(0);}
		
	}
	
	// table 기록
	public static void setting(String filename, String pagenum, String textnum, String textline, String text) {
		
		DefaultTableModel model = (DefaultTableModel)table.getModel();
		
		String[] record = new String[5];
		record[0] = filename;
		record[1] = pagenum;
		record[2] = textnum;
		record[3] = textline;
		record[4] = text;
		
		model.addRow(record);
		
	}
	
}

