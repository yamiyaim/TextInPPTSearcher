import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import com.spire.presentation.IAutoShape;
import com.spire.presentation.ISlide;
import com.spire.presentation.ParagraphEx;
import com.spire.presentation.Presentation;

import java.io.File;
import java.io.FilenameFilter;

public class FileDialog {

	public static void main(String[] args) throws Exception {
		String pathname = new String();
		
		// path고정
		pathname = "C:\\Users\\gmlal\\Desktop\\武逆";
		
		// 검색어
		Scanner sc = new Scanner(System.in);
		System.out.println("word input :");
		String inputText = sc.nextLine();
		sc.close();

		// 해당폴더에있는 전체 ppt파일을 대상으로 타겟워드 검색
		File dir = new File(pathname);
		File[] pptxfiles = dir.listFiles(new FilenameFilter() {

			@Override
			public boolean accept(File dir, String name) {
				return name.endsWith("pptx");
			}
		});
		for (File file : pptxfiles) {
			try {
				List<String> txt = Sercher(GetText(file), inputText);
				for (String text : txt) {
					Log.TraceLog(text);
					System.out.println(text);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	//pptx에서 text파일 추출
	public static List<String> GetText(File file) throws Exception {
		int pagenum = 0;
		int textnum = 0;
		Presentation ppt = new Presentation();
		ppt.loadFromFile(file.toString());
		List<String> text = new ArrayList<>();
		for (Object slide : ppt.getSlides()) {
			pagenum++;
			textnum = 0;
			for (Object shape : ((ISlide) slide).getShapes()) {
				textnum++;
				if (shape instanceof IAutoShape) {
					for (Object tp : ((IAutoShape) shape).getTextFrame().getParagraphs()) {
						text.add(file.getName() + " " + pagenum + "page " + textnum + "text " + ((ParagraphEx) tp).getText());
					}
				}
			}
		}
		return text;
	}

	//textList에서 target을 갖는 text만을 리스트로 반환
	public static List<String> Sercher(List<String> text, String target) {
		List<String> targetintext = new ArrayList<>();
		for (String txt : text) {
			if (txt.contains(target)) {
				targetintext.add(txt);
			}
		}
		return targetintext;
	}

}
