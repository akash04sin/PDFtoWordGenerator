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

public class CopyOfGenerateWordDocument_Testcode {

	public static void generateWord(ExcelSheetDetails sheetDetails, PDFFileDetails pdfFileDetails) throws Exception{
		
		Docx docx = new Docx("D:/innovation/ServerConfigGeneration/Resources/Server_Doc.docx");
		
		//Converting Docx to XWPFDocument
		XWPFDocument document = docx.getXWPFDocument();
		 boolean weMustCommitTableRows = false;
		//Get the table
		XWPFTable table = document.getTableArray(0);
		 XWPFTableRow sourceTableRow = table.getRow(7);
		 //table.removeRow(table.getNumberOfRows() - 1);  
		 XWPFTableRow newRow3 = null;
		 for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){
				String flawIdtext = pdfFileDetails.getFlawDetailArray().get(i).getFlawID();
				 String flawDesctext = pdfFileDetails.getFlawDetailArray().get(i).getFlawName();
				 newRow3 = insertNewTableRow(sourceTableRow, table.getNumberOfRows());
				 int j = 1;
				  newRow3.setCantSplitRow(true);  int row = 0;
				  for (XWPFTableCell cell : newRow3.getTableCells()) {
					 System.out.println("celltext : " + cell.getText());
					// cell.setText(text);
					  //cell.setText(flawIdtext+":"  +flawDesctext);
				  /* for (XWPFParagraph paragraph : cell.getParagraphs()) {
				    for (XWPFRun run : paragraph.getRuns()) {
				     run.setText(flawIdtext +flawDesctext);
				    }
				   }*/
					//row++;
					     XWPFParagraph paragraph = cell.addParagraph();
				            XWPFRun run = paragraph.createRun();
				            run.setText("#{FlawId"+ i + "} : #{Flawdescription" +i+ "}",0);
				            run.addBreak();
				            run.addBreak();
				            run.setText("#{Solution" +i+ "}",1);
				           // run.setText(pdfFileDetails.getFlawDetailArray().get(i).getSolution());
					/*for (XWPFParagraph p : cell.getParagraphs()) {
						  //p.getDocument().setParagraph(paragraph, pos);
						  System.out.println("paragrap text : " + p.getText());
						  String ptext = p.getText();
						  String contenttoReplace = "#{FlawId} : #{Flawdescription}";
						  String contentreplaced="#{FlawId"+ row + "} : #{Flawdescription" +row+ "}";
						  if (ptext != null && ptext.contains(contenttoReplace)){
							 // ptext = ptext.replace(contenttoReplace, contentreplaced);
							  p.getText().replace(contenttoReplace, contentreplaced);
						  }
				            for (XWPFRun r : p.getRuns()) {
				              String text = r.getText(0);
				              if (text != null && text.contains("#{FlawId}")) {
                                  text = text.replace("#{FlawId}", "#{FlawId"+ row + "}");
                                  r.setText(text, 0);
                              }
				            }
					  }*/
				  }
				 
				System.out.println(" number of rows "+table.getNumberOfRows());
				weMustCommitTableRows = true;
				if (weMustCommitTableRows) commitTableRows(table);
			}
		 docx.save("D:/innovation/ServerConfigGeneration/Server_Doc_temp.docx");
		  // now changing something in that new row:
		  
		/*for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){
			table.removeRow(table.getNumberOfRows() - 1);  
			XWPFTableRow lastRow = table.getRows().get(table.getNumberOfRows()-1);
			System.out.println("last row "+table.getNumberOfRows());
			
			CTRow ctrow;
			
			try {
				ctrow = CTRow.Factory.parse(lastRow.getCtRow().newInputStream());
				XWPFTableRow newRow = new XWPFTableRow(ctrow, table);
				
				newRow.getCell(0).setText(" " + "#{FlawId"+i+"}"+ " " +"#{Flawdescription"+i+"}");
				newRow.setCantSplitRow(true);
				table.addRow(newRow );
				
				//System.out.println(table.toString());
			} catch (XmlException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
    	
		}*/	
		//docx.save("D:/innovation/ServerConfigGeneration/Resources/Server_Doc_temp.docx");
		
		//Docx docx = new Docx("D:/innovation/ServerConfigGeneration/Resources/Server_Doc.docx");
		
		Docx docx1 = new Docx("D:/innovation/ServerConfigGeneration/Server_Doc_temp.docx");
		docx1.setVariablePattern(new VariablePattern("#{", "}"));

		
		// preparing variables
		Variables variables = new Variables();
		variables.addTextVariable(new TextVariable("#{AppName}",pdfFileDetails.getProjectName()));
		variables.addTextVariable(new TextVariable("#{EAICode}",pdfFileDetails.getEaicode()));
		variables.addTextVariable(new TextVariable("#{Environment}", sheetDetails.getEnvironment()));
		variables.addTextVariable(new TextVariable("#{AppURL}", pdfFileDetails.getUrlTested()));
		variables.addTextVariable(new TextVariable("#{WebServer}", sheetDetails.getWebServerName()));
		variables.addTextVariable(new TextVariable("#{AppServer}", sheetDetails.getAppServerName()));
	
			/*XWPFParagraph paragraph = document.createParagraph();
        
	      //Set bottom border to paragraph
	      paragraph.setBorderBottom(Borders.BASIC_BLACK_DASHES);
	        
	      //Set left border to paragraph
	      paragraph.setBorderLeft(Borders.BASIC_BLACK_DASHES);
	        
	      //Set right border to paragraph
	      paragraph.setBorderRight(Borders.BASIC_BLACK_DASHES);
	        
	      //Set top border to paragraph
	      paragraph.setBorderTop(Borders.BASIC_BLACK_DASHES);
	      for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){   
	      XWPFRun run = paragraph.createRun();
	         run.setText("#{FlawId"+i+"}" +  "#{Flawdescription"+i+"}");
	      }*/
		/*for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){  
			XWPFTableRow lastRow = table.getRows().get(table.getNumberOfRows() - 1);
			table.removeRow(table.getNumberOfRows() - 1);  
			XWPFTableRow newRow = table.createRow(); 
			 newRow.setCantSplitRow(true);
			    String flawIdtext = pdfFileDetails.getFlawDetailArray().get(i).getFlawID();
			    String flawDesctext = pdfFileDetails.getFlawDetailArray().get(i).getFlawName();
			    
			   
			    XWPFTableCell cell = newRow.getCell(0);
			        if (cell != null) {
			        	//cell.g 
			        	XWPFParagraph paragraph = document.createParagraph();
						XWPFRun run = paragraph.createRun();
			        	cell.setText(flawIdtext + ":" + flawDesctext);
			        	
			        }
			    }*/
	/*	for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){
			String flawIdtext = pdfFileDetails.getFlawDetailArray().get(i).getFlawID();
			 String flawDesctext = pdfFileDetails.getFlawDetailArray().get(i).getFlawName();
		}*/
	   
	/*	Map<String,String> repl1 = new HashMap<String, String>();
		  repl1.put("#{FlawId}", "l1");
		  repl1.put("#{Flawdescription}", "ASDASDASDASD");
		 
		  
		  Map<String,String> repl2 = new HashMap<String,String>();
		  repl1.put("#{FlawId}", "M1");
		  repl1.put("#{Flawdescription}", "NMBMBKJ");
		  try {
			replaceTable(new String[]{"#{FlawId}","#{Flawdescription}"}, Arrays.asList(repl1,repl2), getTemplate("D:/innovation/ServerConfigGeneration/Resources/Server_Doc.docx"));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JAXBException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		*/
		int rowId= 0;
		for(int i =0 ;i<pdfFileDetails.getFlawDetailArray().size();i++){
			System.out.println("PDF File Details  "+pdfFileDetails.getFlawDetailArray().get(i).getFlawID());
			variables.addTextVariable(new TextVariable("#{FlawId"+i+"}", pdfFileDetails.getFlawDetailArray().get(i).getFlawID()));
			variables.addTextVariable(new TextVariable("#{Flawdescription"+i+"}", pdfFileDetails.getFlawDetailArray().get(i).getFlawName()));
			variables.addTextVariable(new TextVariable("#{Solution"+i+"}", pdfFileDetails.getFlawDetailArray().get(i).getSolution()));
		}

		// fill template
		docx1.fillTemplate(variables);
		
		String appName = pdfFileDetails.getProjectName();
		String env = sheetDetails.getEnvironment();
		Path jarDir = Paths.get("");
		System.out.println("FilePath  "+jarDir.toAbsolutePath());
		String currentDirt=jarDir.toAbsolutePath()+"\\";      
		// save filled .docx file
		String outputDocFileName=currentDirt+"Server_Settings_"+appName + "_"+ env +".docx";
		docx1.save(outputDocFileName);
		File filetodelete = new File("D:/innovation/ServerConfigGeneration/Server_Doc_temp.docx");

        if(filetodelete.delete()){
            System.out.println(filetodelete.getName() + " is deleted!");
        }else{
            System.out.println("Delete has failed.");

        }

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
	
	/*private static WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {
		  WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
		  return template;
		 }
	
	private static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements) {
		  Tr workingRow = (Tr) XmlUtils.deepCopy(templateRow);
		  List textElements = getAllElementFromObject(workingRow, Text.class);
		  for (Object object : textElements) {
		   Text text = (Text) object;
		   String replacementValue = (String) replacements.get(text.getValue());
		   if (replacementValue != null)
		    text.setValue(replacementValue);
		  }
		  
		  reviewtable.getContent().add(workingRow);
		 }
	private static Tbl getTemplateTable(List<Object> tables, String templateKey) throws Docx4JException, JAXBException {
		  for (Iterator<Object> iterator = tables.iterator(); iterator.hasNext();) {
		   Object tbl = iterator.next();
		   List<?> textElements = getAllElementFromObject(tbl, Text.class);
		   for (Object text : textElements) {
		    Text textElement = (Text) text;
		    if (textElement.getValue() != null && textElement.getValue().equals(templateKey))
		     return (Tbl) tbl;
		   }
		  }
		  return null;
		 }
	private static void replaceTable(String[] placeholders, List<Map<String, String>> textToAdd,
			   WordprocessingMLPackage template) throws Docx4JException, JAXBException {
			  List<Object> tables = getAllElementFromObject(template.getMainDocumentPart(), Tbl.class);
			  
			  // 1. find the table
			  Tbl tempTable = getTemplateTable(tables, placeholders[0]);
			  List<Object> rows = getAllElementFromObject(tempTable, Tr.class);
			  
			  // first row is header, second row is content
			  if (rows.size() == 7) {
			   // this is our template row
			   Tr templateRow = (Tr) rows.get(7);
			  
			   for (Map<String, String> replacements : textToAdd) {
			    // 2 and 3 are done in this method
			    addRowToTable(tempTable, templateRow, replacements);
			   }
			  
			   // 4. remove the template row
			   tempTable.getContent().remove(templateRow);
			  }
			 }
	private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
		  List<Object> result = new ArrayList<Object>();
		  if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();
		  
		  if (obj.getClass().equals(toSearch))
		   result.add(obj);
		  else if (obj instanceof ContentAccessor) {
		   List<?> children = ((ContentAccessor) obj).getContent();
		   for (Object child : children) {
		    result.addAll(getAllElementFromObject(child, toSearch));
		   }
		  
		  }
		  return result;
		 }*/

	}

