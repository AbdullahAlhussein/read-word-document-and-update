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
	* @author  Abdullah AlHussein
	* 
	*/
	
	 public static void main(String[] args) throws IOException {

		 ReadAndUpdateDocx obj = new ReadAndUpdateDocx();

	        obj.updateDocument(
	                  "./input.docx",
	                  "./output.docx",    //Example input Language LTR
	                  "STC",
	                  "محمد ناصر احمد",
	                  "2020/1/15",
	                  "مطور برامج",
	                  "20000",
	                  "سعد الرماني");
	 }
//		     obj.updateDocument(
//		             "./input.docx",     //Example input Language RTL
//		             "./output2.docx",       
//		             "STC",
//		             "Mohammed ahmed alali",
//		             "2020/1/15",
//		             "software Developer",
//		             "20000",
//		             "Ali nasser");
//		}

	    private void updateDocument(String input, String output, String companyName, String fullName, String startDate, String jobNmae, String salary, String director)
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
	                        docText = docText.replace("${jobName}", jobNmae);
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