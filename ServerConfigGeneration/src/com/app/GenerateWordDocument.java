package com.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import com.app.pojo.ExcelSheetDetails;
import com.app.pojo.PDFFileDetails;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

public class GenerateWordDocument {

	public static void generateWord(ExcelSheetDetails sheetDetails, PDFFileDetails pdfFileDetails) throws Exception{
		
		Path jarDir = Paths.get("");
		String currentDirt = jarDir.toAbsolutePath() + "\\";
		// File folder = new File(currentDirt);
		// Docx docx = new
		// Docx("C:/BACKUP/Learning/ServerConfigGeneration/Resources/Server_Doc.docx");
		Docx docx = new Docx(currentDirt + "Resources/Server_Doc.docx");
			 

		
		//Converting Docx to XWPFDocument
		XWPFDocument document = docx.getXWPFDocument();
		 boolean weMustCommitTableRows = false;
		//Get the table
		XWPFTable table = document.getTableArray(0);
		 XWPFTableRow sourceTableRow = table.getRow(7);//Server_Doc.docx last row is empty Row for flawId and Description

		 XWPFTableRow newRow3 = null;
		 for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){
				String flawIdtext = pdfFileDetails.getFlawDetailArray().get(i).getFlawID();
				String flawDesctext = pdfFileDetails.getFlawDetailArray().get(i).getFlawName();
				 newRow3 = insertNewTableRow(sourceTableRow, table.getNumberOfRows());
				 newRow3.setCantSplitRow(true);  
				  for (XWPFTableCell cell : newRow3.getTableCells()) {
					 System.out.println("celltext : " + cell.getText());
					     XWPFParagraph paragraph = cell.addParagraph();
				            XWPFRun run = paragraph.createRun();
				          
				            run.setFontSize(12);
				            run.setBold(true);
				           // run.setText("#{FlawId"+ i + "} : #{Flawdescription" +i+ "}",0);
				            run.setText(flawIdtext + " : " +  flawDesctext,0);
				            run.addBreak();
				            String data = pdfFileDetails.getFlawDetailArray().get(i).getSolution();
				            if (data.contains("\n")) {
				                String[] lines = data.split("\n");
				                XWPFRun run1 = paragraph.createRun();
				                //run.setText(lines[0], 2); // set first line into XWPFRun
				                System.out.println(lines.length);
				                for(int j=1;j<lines.length;j++){
				                    // add break and insert new text
				                    run1.addBreak();
				                    
				                    run1.setBold(false);
				                    run1.setFontSize(10);
				                    run1.setText(lines[j-1]);
				           
				                }
				            } /*else {
				                run.setText(data, 0);
				            }*/
				            //run.setText(pdfFileDetails.getFlawDetailArray().get(i).getSolution());
				
				  }
				 
				System.out.println(" number of rows "+table.getNumberOfRows());
				weMustCommitTableRows = true;
				if (weMustCommitTableRows) commitTableRows(table);
			}
		 	//docx.save(currentDirt+"Resources/Server_Doc_temp.docx");

			//Reading the temp Docx for replacing values
			//Docx docx1 = new Docx(currentDirt+"Resources/Server_Doc_temp.docx"); 
		
		docx.setVariablePattern(new VariablePattern("#{", "}"));
		
		// preparing variables
		Variables variables = new Variables();
		variables.addTextVariable(new TextVariable("#{AppName}",pdfFileDetails.getProjectName()));
		variables.addTextVariable(new TextVariable("#{EAICode}",pdfFileDetails.getEaicode()));
		variables.addTextVariable(new TextVariable("#{Environment}", sheetDetails.getEnvironment()));
		variables.addTextVariable(new TextVariable("#{AppURL}", pdfFileDetails.getUrlTested()));
		variables.addTextVariable(new TextVariable("#{WebServer}", sheetDetails.getWebServerName()));
		variables.addTextVariable(new TextVariable("#{AppServer}", sheetDetails.getAppServerName()));
	
	/*	for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){
			System.out.println("PDF File Details  "+pdfFileDetails.getFlawDetailArray().get(i).getFlawID());
			variables.addTextVariable(new TextVariable("#{FlawId"+i+"}", pdfFileDetails.getFlawDetailArray().get(i).getFlawID()));
			variables.addTextVariable(new TextVariable("#{Flawdescription"+i+"}", pdfFileDetails.getFlawDetailArray().get(i).getFlawName()));
			//variables.addTextVariable(new TextVariable("#{Solution"+i+"}", "test solution"));
			System.out.println("solution to be added in variable  "+pdfFileDetails.getFlawDetailArray().get(i).getSolution());
		}*/

		// fill template
		docx.fillTemplate(variables);
		
		String appName = pdfFileDetails.getProjectName();
		String env = sheetDetails.getEnvironment();
		//Path jarDir = Paths.get("");
		System.out.println("FilePath  "+jarDir.toAbsolutePath());
		//String currentDirt=jarDir.toAbsolutePath()+"\\";      
		// save filled .docx file
		String outputDocFileName=currentDirt+"Server_Settings_"+appName + "_"+ env +".docx";
		docx.save(outputDocFileName);
	

	}
	static void commitTableRows(XWPFTable table) {
		  int rowNr = 0;
		  for (XWPFTableRow tableRow : table.getRows()) {
		   table.getCTTbl().setTrArray(rowNr++, tableRow.getCtRow());
		  }
		 }
	 static XWPFTableRow insertNewTableRow(XWPFTableRow sourceTableRow, int pos) throws Exception {
		  
		  XWPFTable table = sourceTableRow.getTable();
		 // table.removeRow(table.getNumberOfRows() - i);
		  CTRow newCTRrow = CTRow.Factory.parse(sourceTableRow.getCtRow().newInputStream());
		  XWPFTableRow tableRow = new XWPFTableRow(newCTRrow, table);
		  table.addRow(tableRow, pos);
		
		  return tableRow;
		 }
	

	}

