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

public class TextSearcherinPPT {

	public static void main(String[] args) throws Exception {
		FileSearcher js =new FileSearcher();
	}

}

class FileSearcher extends JFrame implements ActionListener {
	
	String header[] = {"file name", "page", "textbox", "textline", "text"};

	File dir;
	File[] pptxfiles;
	
	JTextField search;
	JFileChooser jfc = new JFileChooser();
	
	JButton btnOpen = new JButton("select folder");
	JButton btnSearch = new JButton("search start");
	JButton btnLogging = new JButton("logging to text");
	JPanel centerPanel = new JPanel();
	JLabel dirpath = new JLabel(" ");
	
	JScrollPane scrolledTable;
	static JTable table;
	
	public FileSearcher() throws Exception {
		this.setTitle("TextSearcherinPPT");
		this.setSize(500, 500);
		this.setLocation(100, 100);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setLayout(new BorderLayout());
		this.panelinit();
		this.tableinit();
		this.set();
		this.setVisible(true);
	}
	
	public void panelinit() {
		JPanel topPanel = new JPanel();
		JPanel bottomPanel = new JPanel();
		JLabel searchlabel = new JLabel("search target : ");
		search = new JTextField(10);

		topPanel.add(btnOpen);
		topPanel.add(dirpath);
		centerPanel.add(searchlabel);
		centerPanel.add(search);
		centerPanel.add(btnSearch);

		add(topPanel, BorderLayout.NORTH);		

		bottomPanel.add(btnLogging);
		add(bottomPanel, BorderLayout.SOUTH);		
	}
	public void tableinit() {
		
		DefaultTableModel model = new DefaultTableModel(header,0);
		
		table = new JTable(model);
		table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
		scrolledTable = new JScrollPane(table);
		scrolledTable.setBorder(BorderFactory.createEmptyBorder(0,10,0,10));
		
		centerPanel.add(scrolledTable);
		add(centerPanel,BorderLayout.CENTER);
	}
	
	public void set() {
		btnOpen.addActionListener(this);
		btnSearch.addActionListener(this);
		btnLogging.addActionListener(this);
		jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		jfc.setMultiSelectionEnabled(false);
	}
	
	//버튼이 눌렸을 경우
	public void actionPerformed(ActionEvent arg0) {
		if (arg0.getSource() == btnOpen) {
			if (jfc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
				dirpath.setText("filepath : " + jfc.getSelectedFile().toString());
				dir = jfc.getSelectedFile();
				pptxfiles = dir.listFiles(new FilenameFilter() {
					
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
			
		} else if (arg0.getSource() == btnLogging) {
			DefaultTableModel model = (DefaultTableModel)table.getModel();
			StringBuilder log = new StringBuilder();
			for (int i=0;i<model.getRowCount();i++) {
				for (int j=0;j<5;j++) {
					log.append(header[j]);
					log.append(" : ");
					log.append(model.getValueAt(i, j).toString());
					log.append(", ");
				}
				int loglength = log.length();
				if (loglength>0) {log.delete(loglength-2, loglength);}
				Log.LogtoTxt(dir.getPath(), log.toString());
				log.delete(0, log.length());
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

