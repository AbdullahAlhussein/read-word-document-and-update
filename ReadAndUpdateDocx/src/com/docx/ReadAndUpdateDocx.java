package com.docx;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ReadAndUpdateDocx {

	/**
	* @author AlHussein
	* 
	*/
	
	 public static void main(String[] args) throws IOException {

		 ReadAndUpdateDocx obj = new ReadAndUpdateDocx();

	        obj.updateDocument(		//Example input Language LTR
	                  "./input.docx",
	                  "./output.docx",    
	                  "شركة الجبر",
	                  "محمد ناصر احمد",
	                  "2018/1/15",
	                  "محاسب",
	                  "11500",
	                  "سعد الرماني");
	 }
//		     obj.updateDocument(	//Example input Language RTL
//		             "./input.docx",     
//		             "./output2.docx",       
//		             "Algebra company",
//		             "Mohammed ahmed alali",
//		             "2018/1/15",
//		             "accountant",
//		             "11500",
//		             "Ali nasser");
//		}

	    private void updateDocument(String input, String output, String companyName, String fullName, String startDate, String jobTitles, String salary, String director)
	            throws IOException {

	            try (XWPFDocument doc = new XWPFDocument(
	                    Files.newInputStream(Paths.get(input)))
	            ) {

	                List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
	                //Iterate over paragraph list and check for the replaceable text in each paragraph
	                for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
	                    for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
	                        String docText = xwpfRun.getText(0);
	                        //replacement and setting position
	                        docText = docText.replace("${companyName}", companyName);
	                        docText = docText.replace("${fullName}", fullName);
	                        docText = docText.replace("${startDate}", startDate);
	                        docText = docText.replace("${jobTitles}", jobNmae);
	                        docText = docText.replace("${salary}", salary);
	                        docText = docText.replace("${director}", director);
	                        xwpfRun.setText(docText, 0);
	                    }
	                }

	            // save the docs
	            try (FileOutputStream out = new FileOutputStream(output)) {
	                doc.write(out);
	            }

	        }

	    }
	}
