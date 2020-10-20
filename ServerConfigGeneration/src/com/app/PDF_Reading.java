package com.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.app.pojo.FlawDetail;
import com.app.pojo.PDFFileDetails;

public class PDF_Reading {

	static Properties prop = new Properties();
	static List<String> serverConfigFlaw = new ArrayList<String>();
	static List<String> flawDetail = new ArrayList<String>();
	public static void main(String[] args) throws IOException {

		/*String solution = "A) Server Version
							i) Open configuration file and add the following if missing
								DEV- /metlife/runtime/content/cppDEV1/salespilot/run/conf/httpd.conf
								ServerTokens Prod
								ServerSignature Off
								AddServerHeader off
								Click on Apply.
						B) X-PoweredBy
							i) Go to Application servers > [Application JVM]> Web Container Settings >Web container > Custom properties 
							ii) Click on New and add the below property 
								Name: com.ibm.ws.webcontainer.disablexPoweredBy
								Value: true
							iii) Click on Save.";
*/		Path jarDir = Paths.get("");
		System.out.println("FilePath  " + jarDir.toAbsolutePath());
		String currentDirt = jarDir.toAbsolutePath() + "\\";
		File folder = new File(currentDirt);
		File[] listOfFiles = folder.listFiles();
		String excelSolutionFilePath="";
		String excelFilefilePath = "";
		for (int i = 0; i < listOfFiles.length; i++) {
			String listOfFile = listOfFiles[i].getName();
			if (listOfFile.endsWith(".properties")) {
				FileInputStream fis = new FileInputStream(listOfFile);
				prop.load(fis);
				String serverCon = prop.getProperty("serverConfigFlaw");
				serverConfigFlaw = Arrays.asList(serverCon.split(";"));
			}
			if (listOfFile.endsWith(".xlsx")) {
				if(listOfFile.contains("Solution")){
					excelSolutionFilePath = listOfFiles[i].getName();
					}
				else{
				excelFilefilePath = listOfFiles[i].getName();
				}
			}
			/*if (listOfFile.endsWith(".pdf")) {
				excelFilefilePath = listOfFiles[i].getName();
			}*/
		}
		
		
		for (int i = 0; i < listOfFiles.length; i++) {
			String pdfFile = listOfFiles[i].getName();
			//System.out.println(pdfFile);
			if (pdfFile.startsWith("MET") && pdfFile.endsWith(".pdf")) {
				String[] pdfFileNameArr = pdfFile.split("-");
				int pwdyr = 2017;
				boolean wrongPwd = true;
				while (wrongPwd) {
					try {
						//System.out.println(pwdyr + "pwdyr");
						PDDocument doc = PDDocument.load(
								listOfFiles[i].getAbsoluteFile(), "$#" + pwdyr
										+ "_" + pdfFileNameArr[2] + "#$");
						PDFTextStripper pdfStripper = new PDFTextStripper();
						pdfStripper.setEndPage(2);
						String b = pdfStripper.getText(doc).trim();
						//System.out.println("Text in PDF\n---------------------------------");
						String[] lines = b.split("\\r?\\n");
						PDFFileDetails pdffiledtobj=new PDFFileDetails();
						//System.out.println(pdffiledtobj);
						boolean  serverCof= findSeverCofigFlaw(lines);
						
						if (serverCof) {
							List<FlawDetail> flawDetailList=new ArrayList<FlawDetail>();
							for (int j = 0; j < lines.length; j++) {//loop for pdf line by line
								// System.out.println(lines[j]);
								// +"+_+++++++++++++");
								if (lines[j].contains("Project Name")) {
									String projectName = lines[j].substring(12)
											.trim();
									if(projectName.contains("Company")){
									    String projectNameSplit[]=	projectName.split("Company");
										 projectName = projectNameSplit[0].trim();
									}
									pdffiledtobj.setProjectName(projectName);
									//System.out.println(projectName+"projectName");
								} else if (lines[j].contains("Issued")) {
									String issuedDate = lines[j + 1].trim();
									pdffiledtobj.setIssuedDate(issuedDate);
								} else if (lines[j].contains("Project #")) {
									String projectCode = lines[j].substring(9).trim();
									pdffiledtobj.setProjectCode(projectCode);
								} else if (lines[j].contains("EAI #")) {
									String eaicode = lines[j].substring(5).trim();
									pdffiledtobj.setEaicode(eaicode);

								} else if (lines[j].contains("URL Tested")) {
									String urlTested = lines[j].substring(10).trim();
									pdffiledtobj.setUrlTested(urlTested);
								}else {
									Iterator itr = serverConfigFlaw.iterator();
									while (itr.hasNext()) {//iterator for property file 
										FlawDetail flawdtobj=new FlawDetail();
										String coparingFlaw=(String) itr.next();
										if (lines[j].contains(coparingFlaw) && !(lines[j].contains("Pre-Authentication"))) {
											//String flawID=;
											flawdtobj.setFlawName(coparingFlaw);
											String[] coparingFlawSplit= lines[j].split(coparingFlaw.trim());
											String[] severitySplit2=(coparingFlawSplit[1].trim()).split(" ");
											flawdtobj.setFlawID(coparingFlawSplit[0]);
											flawdtobj.setSeverity(severitySplit2[0]);
											flawdtobj.setStatus(severitySplit2[severitySplit2.length-2]);
											flawdtobj.setSolution(solution(excelSolutionFilePath,coparingFlaw));
											flawDetailList.add(flawdtobj);
											pdffiledtobj.setFlawDetailArray(flawDetailList);
										}
									}
									pdffiledtobj.setFlawDetailArray(flawDetailList);
								}
							}
							//call excel method
							//String filePath = "D:\innovation\ServerConfigGeneration\Webserver_AppserverDetails1.xlsx";
							ReadExcel.readExcel(excelFilefilePath,pdffiledtobj);
							/*File filetodelete = new File("D:/innovation/ServerConfigGeneration/Server_Doc_temp.docx");

					        if(filetodelete.delete()){
					            System.out.println(filetodelete.getName() + " is deleted!");
					        }else{
					            System.out.println("Delete has failed.");

					        }*/
							//deleting temp serverdoc file
						
							
							//System.out.println(pdffiledtobj);
						}
						
						doc.close();
						
						wrongPwd = false;
					} catch (Exception e) {
						System.out.println("wrong Password"+e);
						pwdyr++;
						//System.exit(0);
					}
					

				}
			}

		}

	}

	public static boolean findSeverCofigFlaw(String lines[]) {
		
		for (int j = lines.length - 1; j >= 0; j--) {
			Iterator itr = serverConfigFlaw.iterator();
			while (itr.hasNext()) {
			
					if (lines[j].contains((String) itr.next()) && !(lines[j].contains("Pre-Authentication"))) {
					return true;
				}
			}
		}
		return false;
	}
	public static String solution(String excelSolutionFilePath, String coparingFlaw){
		DataFormatter dataFormatter = new DataFormatter();
		String s = "";
		try {
			Workbook workbook = WorkbookFactory.create(new File(excelSolutionFilePath));
			Sheet sheet = workbook.getSheetAt(0);
			int rowsCount = sheet.getPhysicalNumberOfRows();
			for(int i=1;i<rowsCount;i++){
				Row row = sheet.getRow(i);
				if(coparingFlaw.equals(dataFormatter.formatCellValue(row.getCell(0)))){
					
					 s=dataFormatter.formatCellValue(row.getCell(1));
					System.out.println(s);
					return s;
				}
			}
		} catch (EncryptedDocumentException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return s;
		
	}
}
