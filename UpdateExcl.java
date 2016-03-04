package apachaePOI;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.Date;
import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.util.CellReference;

//import org.apache.poi.xssf.usermodel.XSSFSheet;

public class UpdateExcl {

	static int lastrow;
	static int lastRowYesterday;

	public int getLastRownum(int sheetNum,String fileName) throws IOException {

		try {
			FileInputStream file = new FileInputStream(new File(
					fileName));
			// /Users/droy2/Desktop/OrderImportExceptionReport
			// /Users/droy2/Documents/ellipse/Test1.xlsx
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			lastrow = sheet.getLastRowNum();

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return lastrow;

	}
	
	
	public int getYesterdayLastRow(int sheetNum , String updateFlag,int setLastRow,String fileName) throws IOException {
       
		try {
			FileInputStream file = new FileInputStream(new File(
					fileName));
			// /Users/droy2/Desktop/OrderImportExceptionReport
			// /Users/droy2/Documents/ellipse/Test1.xlsx
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			XSSFRow row ;
			row = sheet.getRow(1);
			row.getCell(0).setCellType(Cell.CELL_TYPE_NUMERIC);
			
			if(updateFlag =="Y"){
				row.getCell(0).setCellType(Cell.CELL_TYPE_NUMERIC);
				row.getCell(0).setCellValue((double)setLastRow+1);
			
				FileOutputStream outFile = new FileOutputStream(new File(fileName));
				workbook.write(outFile);
				outFile.close();
				file.close();
				
			}
			lastRowYesterday =  (int)row.getCell(0).getNumericCellValue();
			

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return lastRowYesterday+1;

	}
	
	public int getSiebelYesterdayLastRow(int sheetNum , String updateFlag,int setLastRow,String fileName) throws IOException {
	       
		try {
			FileInputStream file = new FileInputStream(new File(
					fileName));
			// /Users/droy2/Desktop/OrderImportExceptionReport
			// /Users/droy2/Documents/ellipse/Test1.xlsx
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			XSSFRow row ;
			int color;
			row = sheet.getRow(1);
			row.getCell(3).setCellType(Cell.CELL_TYPE_NUMERIC);
			
			if(updateFlag =="Y"){
				row.getCell(3).setCellType(Cell.CELL_TYPE_NUMERIC);
				row.getCell(3).setCellValue((double)setLastRow+1);
				
				color = (int)row.getCell(2).getNumericCellValue();
				   
				   System.out.println("Last COlor was" +color );
				   
				   if (color ==1){
					   row.getCell(2).setCellValue((int) 2);
					   
				   }
				   if (color ==2){
					   row.getCell(2).setCellValue((int) 1);
					   
				   }
				   
				    FileOutputStream outFile = new FileOutputStream(new File(fileName));
					workbook.write(outFile);
					outFile.close();
					file.close();
			
				
				
			}
			
		   
		   
			lastRowYesterday =  (int)row.getCell(3).getNumericCellValue();
			

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return lastRowYesterday+1;

	}
	
	

	public void updateSheetRow(int sheetNum, int firstRow, int lastRow,
			String fileName, String value) throws IOException {
		FileInputStream file;
		try {
			file = new FileInputStream(new File(fileName));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			XSSFSheet sheet1 = workbook.getSheetAt(1);
			XSSFRow row,row1;

			for (int i = firstRow - 1; i <= lastRow; i++) {
				row = sheet.getRow(i);
				row1 = sheet1.getRow(i);
				// row.createCell((short) 0).setCellValue("Existing");
				
				

				row.getCell(0).setCellValue(value);
				
				
				
				System.out.println("Siebel Orders"+row.getCell(6).getStringCellValue());

			}
			
			/*for (int i = 13236; i <= 13240; i++) {
				row1 = sheet1.getRow(i);
				// row.createCell((short) 0).setCellValue("Existing");

				
				System.out.println("Siebel Orders Status"+row1.getCell(0).getStringCellValue());

			}*/
			
			
			
			FileOutputStream outFile = new FileOutputStream(new File(fileName));
			workbook.write(outFile);
			outFile.close();
			file.close();

		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

	}

	
	public void updateSiebelFlag(int sheetNum, int firstRow0, int lastRow0,int firstRow1 ,int lastRow1,
			String fileName, String value) throws IOException {
		FileInputStream file;
		try {
			file = new FileInputStream(new File(fileName));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			XSSFSheet sheet1 = workbook.getSheetAt(1);
			XSSFRow row,row1;
			
			
			for (int i = firstRow0-1; i <= lastRow0; i++) {
				row = sheet.getRow(i);
				for (int j = firstRow1-1; j <= lastRow1; j++) {
					row1 = sheet1.getRow(j);
					XSSFRichTextString Str1;
					XSSFRichTextString Str2;
					Str1=row.getCell(6).getRichStringCellValue();
					Str2=row1.getCell(0).getRichStringCellValue();
					
					/////////////////////////////////////////////
					/*switch (row.getCell(6).getCellType()) {
	                case Cell.CELL_TYPE_STRING:
	                    System.out.println("The cell type is String");
	                    break;
	                case Cell.CELL_TYPE_NUMERIC:
	                    if (DateUtil.isCellDateFormatted(row.getCell(6))) {
	                        System.out.println("It is Date Type");
	                    } else {
	                        System.out.println("It is Numeric Type");
	                    }
	                    break;
	                case Cell.CELL_TYPE_BOOL	EAN:
	                    System.out.println("It is boolean");
	                    break;
	                case Cell.CELL_TYPE_FORMULA:
	                    System.out.println("It is formula");
	                    break;
	                default:
	                    System.out.println("Its nothing");
	            }*/
			/////////////////////////////////////////////////////////
				/*	switch (row1.getCell(0).getCellType()) {
	                case Cell.CELL_TYPE_STRING:
	                    System.out.println("The cell type is String");
	                    break;
	                case Cell.CELL_TYPE_NUMERIC:
	                    if (DateUtil.isCellDateFormatted(row1.getCell(0))) {
	                        System.out.println("It is Date Type");
	                    } else {
	                        System.out.println("It is Numeric Type");
	                    }
	                    break;
	                case Cell.CELL_TYPE_BOOLEAN:
	                    System.out.println("It is boolean");
	                    break;
	                case Cell.CELL_TYPE_FORMULA:
	                    System.out.println("It is formula");
	                    break;
	                default:
	                    System.out.println("Its nothing");
	            }*/
		////////////////////////////////////////////////////////////////			
					
					
					
					
					///////////////////////////////////////////////
					
					
					
					if (Str1.getString().equals(Str2.getString())){ 
						System.out.println("The Value of Order in SHeet0 "+ Str1);
						System.out.println("The Value of Order in SHeet1 "+ Str1);
						
						System.out.println("True COndition The cell type is String");
						System.out.println("Setting value to True");
						
						switch (row1.getCell(2).getCellType()) {
		                case Cell.CELL_TYPE_STRING:
		                	row.getCell(20).setCellValue(row1.getCell(2).getStringCellValue());
		                    break;
		                case Cell.CELL_TYPE_NUMERIC:
		                	 System.out.println("Its String");
		                    break;
		                case Cell.CELL_TYPE_BOOLEAN:
		                    row.getCell(20).setCellValue(row1.getCell(2).getBooleanCellValue());
		                    break;
		                case Cell.CELL_TYPE_FORMULA:
		                    System.out.println("It is formula");
		                    break;
		                default:
		                    System.out.println("Its nothing");
		            }
						break;
						
						
						//row.getCell(20).setCellValue("TRUE");
						//break;
						
						//row.getCell(20).setCellValue(row1.getCell(2).getStringCellValue());
						//row.getCell(20).setCellType(Cell.CELL_TYPE_BOOLEAN);
						
					}else {
						System.out.println("The Value of Order in SHeet0 "+ Str1);
						System.out.println("The Value of Order in SHeet1 "+ Str2);
						
						//row.getCell(20).setCellType(Cell.CELL_TYPE_BOOLEAN);
						//row.getCell(20).setCellValue(new XSSFRichTextString("FALSE"));
						//row.getCell(20).getCellType();
						//row.getCell(20).setCellValue("FALSE");
						//row.getCell(20).setCellType(Cell.CELL_TYPE_BOOLEAN);
						
						/*switch (row.getCell(20).getCellType()) {
		                case Cell.CELL_TYPE_STRING:
		                    System.out.println("The cell type is String");
		                    break;
		                case Cell.CELL_TYPE_NUMERIC:
		                    if (DateUtil.isCellDateFormatted(row.getCell(20))) {
		                        System.out.println("It is Date Type");
		                    } else {
		                        System.out.println("It is Numeric Type");
		                    }
		                    break;
		                case Cell.CELL_TYPE_BOOLEAN:
		                    System.out.println("It is boolean");
		                    break;
		                case Cell.CELL_TYPE_FORMULA:
		                    System.out.println("It is formula");
		                    break;
		                default:
		                    System.out.println("Its nothing");
		            }*/
						
						System.out.println("Value is set to False");
					}
						
					
				}
				

			}
			
			
			
			
			
			FileOutputStream outFile = new FileOutputStream(new File(fileName));
			workbook.write(outFile);
			outFile.close();
			file.close();

		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

	}

	
	public void updateSiebelOrders(int sheetNum, int firstRow, int lastRow,
			String fileName) throws IOException {
		FileInputStream file;
		try {
			file = new FileInputStream(new File(fileName));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			XSSFSheet sheet1 = workbook.getSheetAt(1);
			XSSFRow row0, row1;

			for (int i = firstRow; i <= lastRow; i++) {
				row0 = sheet.getRow(i);

				for (int j = firstRow; j <= lastRow; j++) {
					row1 = sheet1.getRow(j);
					if (row0.getCell(20).getRichStringCellValue()
							.equals(row1.getCell(0))) {

						row0.getCell(20).setCellValue(
								row1.getCell(0).getRawValue());
						

					} else {
						row0.getCell(20).setCellValue("FALSETEST");
					}

				}
				FileOutputStream outFile = new FileOutputStream(new File(
						fileName));
				workbook.write(outFile);
				outFile.close();
				file.close();
			}
		}

		catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

	}

	public void Update(int sheetNum, int lastRow, int colSpan, String[] fields,String filename) {
		try {
			//String str1, str2;
			FileInputStream file = new FileInputStream(new File(
					filename));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sheetNum);
			//Updated Oct 8 2013 changed sheet num 4 to 2
			XSSFSheet sheetH = workbook.getSheetAt(2);
			XSSFRow rowH = sheetH.getRow(1);
			
			CreationHelper createHelper = workbook.getCreationHelper();
			// FormulaEvaluator evaluator =
			// workbook.getCreationHelper().createFormulaEvaluator();
			// sheet.createRow(18);
			// System.out.println("The Row Num is "+ lastRow);
			/*
			 * for(int i=0;i<colSpan;i++){
			 * 
			 * System.out.println("Row Value "+lastrow+"Col "+ i+"Value ="+
			 * fields[i]+"Total COlumns "+colSpan); XSSFRow row=
			 * sheet.createRow((short)lastRow); row.createCell((short)
			 * i).setCellValue(fields[i]);
			 * 
			 * 
			 * }
			 */

			XSSFRow row = sheet.createRow((short) lastRow);
			//XSSFRow prevrow = sheet.getRow(lastRow - 1);// The previous row
			row.createCell((short) 0).setCellValue("New");

			// Cell cell = row.getCell(0);
			// cell.setCellFormula("IF(INDIRECT(" + "\"" + "V" + "\""
			// + "&ROW()) = MAX($V:V)," + "\"" + "New" + "\" " + ","
			// + "\"" + "Existing" + " \"" + ")");
			//FormulaEvaluator evaluator = workbook.getCreationHelper()
				//	.createFormulaEvaluator();
			// evaluator.evaluateFormulaCell(cell);
			// System.out.println("The cell type :"+
			// evaluator.evaluateFormulaCell(cell));
			if (colSpan == 3) {
				row.createCell((short) 0).setCellValue(fields[0]);
				row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
				
				row.createCell((short) 1).setCellValue(fields[1]);
				row.createCell((short) 2).setCellValue(fields[2]);
				row.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
			}
			if (colSpan == 2) {
				row.createCell((short) 0).setCellValue(fields[0]);
				
				row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
				
				row.createCell((short) 1).setCellValue(fields[1]);

			}
			if (colSpan == 1) {
				row.createCell((short) 0).setCellValue(fields[0]);
				row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);

			}
			if (colSpan == 23) {
				row.createCell((short) 1).setCellValue(fields[1]);
				// ///////////////////////////////////////////////
				row.createCell((short) 2).setCellValue(fields[2]);
				// Cell cell2 = row.getCell(2);
				// cell2.setCellType(Cell.CELL_TYPE_NUMERIC);
				// evaluator.evaluateFormulaCell(cell2);
				// ////////////////////////////////////////////////
				row.createCell((short) 3).setCellValue(fields[3]);
				row.createCell((short) 4).setCellValue(fields[4]);
				//////////////////////////////////////////////////
				
				Double convertedNumber = Double.valueOf(NVL(fields[5],0));
				row.createCell(5).setCellValue(convertedNumber);
                row.getCell(5).setCellType(Cell.CELL_TYPE_NUMERIC);
				//////////////////////////////////////////////////
				
				row.createCell((short) 6).setCellValue(fields[6]);
				row.createCell((short) 7).setCellValue(fields[7]);
				row.createCell((short) 8).setCellValue(fields[8]);
				row.createCell((short) 9).setCellValue(fields[9]);
				row.createCell((short) 10).setCellValue(fields[10]);
				row.createCell((short) 11).setCellValue(fields[11]);
				row.createCell((short) 12).setCellValue(fields[12]);
				row.createCell((short) 13).setCellValue(fields[13]);
				row.createCell((short) 14).setCellValue(fields[14]);
				row.createCell((short) 15).setCellValue(fields[15]);
				// ///////////////////////////////////////////////////
				 Double convertedNumberQ = Double.valueOf(NVL(fields[16],0));
                 row.createCell(16).setCellValue(convertedNumberQ);
                 row.getCell(16).setCellType(Cell.CELL_TYPE_NUMERIC);
				// /////////////////////////////////////////////////////

				// row.createCell((short) 17).setCellValue(fields[17]);
				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setDataFormat(createHelper.createDataFormat()
						.getFormat("m/d/yy"));
				Cell cell17 = row.createCell(17);
				//Date date = new Date(System.currentTimeMillis());
				System.out.println("Date is " + fields[17]);
				cell17.setCellValue(fields[17]);
				cell17.setCellStyle(cellStyle);
				FormulaEvaluator evaluator17 = workbook.getCreationHelper()
						.createFormulaEvaluator();
				evaluator17.evaluateFormulaCell(cell17);
				// /////////////////////////////////////////////////////////////
				// /////////////////////////////////////////////////////////////

				// Cell cell18 = row.getCell(18);
				// CellStyle cellStyle18 = workbook.createCellStyle();
				// cell18.setCellType(Cell.CELL_TYPE_NUMERIC);
				row.createCell((short) 18).setCellValue(fields[18]);
				//Cell cell18 = row.getCell(18);
				// CellStyle cellStyle18 = workbook.createCellStyle();
				//cell18.setCellType(Cell.CELL_TYPE_NUMERIC);

				// foo = Integer.parseInt(fields[18]);
				// /////////////////////////////////////////////////////
				row.createCell((short) 19).setCellValue(fields[19]);
				// ////////////////////////////////////////////////////
				
				row.createCell((short) 20).setCellValue("FALSE");
				//row.getCell(20).setCellType(Cell.CELL_TYPE_BOOLEAN);
				/*
				 * row.createCell((short) 20)
				 * .setCellValue((fields[20].toString())); Cell cell20 =
				 * row.getCell(20); cell20.setCellType(Cell.CELL_TYPE_BOOLEAN);
				 * // cell20.setCellValue (fields[20].toString()); //
				 * cell20.setCellFormula("IFERROR(VLOOKUP(INDIRECT(" +"\"" +"G"
				 * // +"\"" + "&ROW()),'Orders in Siebel'!A:C,3,FALSE),FALSE)");
				 * // FormulaEvaluator evaluator2 = //
				 * workbook.getCreationHelper().createFormulaEvaluator();
				 * evaluator.evaluateFormulaCell(cell20);
				 * System.out.println("The cell type :" +
				 * evaluator.evaluateFormulaCell(cell20));
				 */
				// //////////////////////////////////////////////////////////////////
				CellStyle style = sheet.getWorkbook().createCellStyle();
				style.setDataFormat((short) BuiltinFormats
						.getBuiltinFormat("m/d/yy"));

				String strDate;
				strDate = "DATE("
						+ fields[21].substring(fields[21].lastIndexOf("/") + 1)
						+ ","
						+ fields[21].substring(0, fields[21].indexOf("/"))
						+ ","
						+ fields[21].substring(fields[21].indexOf("/") + 1,
								fields[21].lastIndexOf("/")) + ")";

				System.out.println("The value of strDate is : " + strDate);

				// ////////////////////////Added////////////////////////////
				row.createCell((short) 21).setCellValue(fields[21]);
				//row.getCell(21).setCellType(Cell.CELL_TYPE_NUMERIC);
				
				
				// //////////////////Cell 22/////////////////////////////////
				
				
				row.createCell((short) 22).setCellValue((int)rowH.getCell(2).getNumericCellValue());
				System.out.println("Color :"+ (int)rowH.getCell(2).getNumericCellValue());
				
				
			}

			if (colSpan == 24) {
				row.createCell((short) 1).setCellValue(fields[1]);
				// ///////////////////////////////////////////////
				row.createCell((short) 2).setCellValue(fields[2]);
				// Cell cell2 = row.getCell(2);
				// cell2.setCellType(Cell.CELL_TYPE_NUMERIC);
				// evaluator.evaluateFormulaCell(cell2);
				// ////////////////////////////////////////////////
				row.createCell((short) 3).setCellValue(fields[3]);
				row.createCell((short) 4).setCellValue(fields[4]);
				
              //////////////////////////////////////////////////
				Double convertedNumber = Double.valueOf(NVL(fields[5],0));
				row.createCell(5).setCellValue(convertedNumber);
                row.getCell(5).setCellType(Cell.CELL_TYPE_NUMERIC);
                
              //////////////////////////////////////////////////
				row.createCell((short) 6).setCellValue(fields[6]);
				row.createCell((short) 7).setCellValue(fields[7]);
				row.createCell((short) 8).setCellValue(fields[8]);
				row.createCell((short) 9).setCellValue(fields[9]);
				row.createCell((short) 10).setCellValue(fields[10]);
				row.createCell((short) 11).setCellValue(fields[11]);
				row.createCell((short) 12).setCellValue(fields[12]);
				row.createCell((short) 13).setCellValue(fields[13]);
				row.createCell((short) 14).setCellValue(fields[14]);
				row.createCell((short) 15).setCellValue(fields[15]);
				// /////////////////////////////////////////////////
				Double convertedNumberQ = Double.valueOf(NVL(fields[16],0));
                row.createCell(16).setCellValue(convertedNumberQ);
                row.getCell(16).setCellType(Cell.CELL_TYPE_NUMERIC);
                
				// /////////////////////////////////////////////////////

				// row.createCell((short) 17).setCellValue(fields[17]);
				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setDataFormat(createHelper.createDataFormat()
						.getFormat("m/d/yy"));
				Cell cell17 = row.createCell(17);
				//Date date = new Date(9);
				System.out.println("Date is " + fields[17]);
				cell17.setCellValue(fields[17]);
				cell17.setCellStyle(cellStyle);
				FormulaEvaluator evaluator17 = workbook.getCreationHelper()
						.createFormulaEvaluator();
				evaluator17.evaluateFormulaCell(cell17);
				// /////////////////////////////////////////////////////////////
				// /////////////////////////////////////////////////////////////
				// Cell cell18 = row.getCell(18);
				// CellStyle cellStyle18 = workbook.createCellStyle();
				// cell18.setCellType(Cell.CELL_TYPE_NUMERIC);
				row.createCell((short) 18).setCellValue(fields[18]);
				//Cell cell18 = row.getCell(18);
				// CellStyle cellStyle18 = workbook.createCellStyle();
				//cell18.setCellType(Cell.CELL_TYPE_NUMERIC);
				// foo = Integer.parseInt(fields[18]);

				// /////////////////////////////////////////////////////
				row.createCell((short) 19).setCellValue(fields[19]);
				// ////////////////////////////////////////////////////
				row.createCell((short) 20).setCellValue("FALSE");
				//row.getCell(20).setCellType(Cell.CELL_TYPE_BOOLEAN);

				/*
				 * row.createCell((short) 20)
				 * .setCellValue((fields[20].toString())); Cell cell20 =
				 * row.getCell(20); cell20.setCellType(Cell.CELL_TYPE_BOOLEAN);
				 * // cell20.setCellValue (fields[20].toString()); //
				 * cell20.setCellFormula("IFERROR(VLOOKUP(INDIRECT(" +"\"" +"G"
				 * // +"\"" + "&ROW()),'Orders in Siebel'!A:C,3,FALSE),FALSE)");
				 * // FormulaEvaluator evaluator2 = //
				 * workbook.getCreationHelper().createFormulaEvaluator();
				 * evaluator.evaluateFormulaCell(cell20); ////tobe
				 * continued///// System.out.println("The cell type :" +
				 * evaluator.evaluateFormulaCell(cell20));
				 */
				// //////////////////////////////////////////////////////////////////
				String strDate;
				strDate = "DATE("
						+ fields[21].substring(fields[21].lastIndexOf("/") + 1)
						+ ","
						+ fields[21].substring(0, fields[21].indexOf("/"))
						+ ","
						+ fields[21].substring(fields[21].indexOf("/") + 1,
								fields[21].lastIndexOf("/")) + ")";

				System.out.println("The value of strDate is : " + strDate);

				

				// ////////////////////////Added////////////////////////////
				row.createCell((short) 21).setCellValue(fields[21]);
				//row.getCell(21).setCellType(Cell.CELL_TYPE_NUMERIC);
				
				
				// //////////////////Cell 22/////////////////////////////////
				
				
				row.createCell((short) 22).setCellValue((int)rowH.getCell(2).getNumericCellValue());
				
				System.out.println("Color :"+ (int)rowH.getCell(2).getNumericCellValue());
				// ///////////////////////////////////////////////////
				row.createCell((short) 23).setCellValue(fields[23]);

			}

			if (colSpan == 25) {
				row.createCell((short) 1).setCellValue(fields[1]);
				// ///////////////////////////////////////////////
				row.createCell((short) 2).setCellValue(fields[2]);
				// Cell cell2 = row.getCell(2);
				// cell2.setCellType(Cell.CELL_TYPE_NUMERIC);
				// evaluator.evaluateFormulaCell(cell2);
				// ////////////////////////////////////////////////
				row.createCell((short) 3).setCellValue(fields[3]);
				row.createCell((short) 4).setCellValue(fields[4]);
                  //////////////////////////////////////////////////
				Double convertedNumber = Double.valueOf(NVL(fields[5],0));
				row.createCell(5).setCellValue(convertedNumber);
                row.getCell(5).setCellType(Cell.CELL_TYPE_NUMERIC);
                 
                 //////////////////////////////////////////////////
				row.createCell((short) 6).setCellValue(fields[6]);
				row.createCell((short) 7).setCellValue(fields[7]);
				row.createCell((short) 8).setCellValue(fields[8]);
				row.createCell((short) 9).setCellValue(fields[9]);
				row.createCell((short) 10).setCellValue(fields[10]);
				row.createCell((short) 11).setCellValue(fields[11]);
				row.createCell((short) 12).setCellValue(fields[12]);
				row.createCell((short) 13).setCellValue(fields[13]);
				row.createCell((short) 14).setCellValue(fields[14]);
				row.createCell((short) 15).setCellValue(fields[15]);
				// //////////////////////////////////////////////////
				Double convertedNumberQ = Double.valueOf(NVL(fields[16],0));
                row.createCell(16).setCellValue(convertedNumberQ);
                row.getCell(16).setCellType(Cell.CELL_TYPE_NUMERIC);
                
				// /////////////////////////////////////////////////////

				// row.createCell((short) 17).setCellValue(fields[17]);
				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setDataFormat(createHelper.createDataFormat()
						.getFormat("m/d/yy"));
				Cell cell17 = row.createCell(17);
				//Date date = new Date(9);
				System.out.println("Date is " + fields[17]);
				cell17.setCellValue(fields[17]);
				cell17.setCellStyle(cellStyle);
				FormulaEvaluator evaluator17 = workbook.getCreationHelper()
						.createFormulaEvaluator();
				evaluator17.evaluateFormulaCell(cell17);
				/*
				 * SimpleDateFormat format1 = new
				 * SimpleDateFormat("MM/dd/yyyy"); SimpleDateFormat format2 =
				 * new SimpleDateFormat("dd-MMM-yy"); Date date1 =
				 * format1.parse("05/01/1999");
				 * System.out.println(format2.format(date));
				 */

				// /////////////////////////////////////////////////////////////

				// Cell cell18 = row.getCell(18);
				// CellStyle cellStyle18 = workbook.createCellStyle();
				// cell18.setCellType(Cell.CELL_TYPE_NUMERIC);
				row.createCell((short) 18).setCellValue(fields[18]);
				//Cell cell18 = row.getCell(18);
				// CellStyle cellStyle18 = workbook.createCellStyle();
				//cell18.setCellType(Cell.CELL_TYPE_NUMERIC);
				// foo = Integer.parseInt(fields[18]);

				// /////////////////////////////////////////////////////
				row.createCell((short) 19).setCellValue(fields[19]);
				// ////////////////////////////////////////////////////
				row.createCell((short) 20).setCellValue("FALSE");
				//row.getCell(20).setCellType(Cell.CELL_TYPE_BOOLEAN);

				/*
				 * row.createCell((short) 20)
				 * .setCellValue((fields[20].toString())); Cell cell20 =
				 * row.getCell(20); cell20.setCellType(Cell.CELL_TYPE_BOOLEAN);
				 * // cell20.setCellValue (fields[20].toString()); //
				 * cell20.setCellFormula("IFERROR(VLOOKUP(INDIRECT(" +"\"" +"G"
				 * // +"\"" + "&ROW()),'Orders in Siebel'!A:C,3,FALSE),FALSE)");
				 * // FormulaEvaluator evaluator2 = //
				 * workbook.getCreationHelper().createFormulaEvaluator();
				 * evaluator.evaluateFormulaCell(cell20);
				 * System.out.println("The cell type :" +
				 * evaluator.evaluateFormulaCell(cell20));
				 */
				// //////////////////////////////////////////////////////////////////
				String strDate;
				strDate = "DATE("
						+ fields[21].substring(fields[21].lastIndexOf("/") + 1)
						+ ","
						+ fields[21].substring(0, fields[21].indexOf("/"))
						+ ","
						+ fields[21].substring(fields[21].indexOf("/") + 1,
								fields[21].lastIndexOf("/")) + ")";

				System.out.println("The value of strDate is : " + strDate);

				// ////////////////////////Added////////////////////////////
				row.createCell((short) 21).setCellValue(fields[21]);
				//row.getCell(21).setCellType(Cell.CELL_TYPE_NUMERIC);
				
				
				// //////////////////Cell 22/////////////////////////////////
				
				
				row.createCell((short) 22).setCellValue((int)rowH.getCell(2).getNumericCellValue());
				System.out.println("Color :"+ (int)rowH.getCell(2).getNumericCellValue());

				// ///////////////////////////////////////////////////

				row.createCell((short) 23).setCellValue(fields[23]);
				row.createCell((short) 24).setCellValue(fields[24]);

			}
			// row.createCell((short) 24).setCellValue(fields[24]);
			// row.createCell((short) 25).setCellValue(fields[25]);

			// Cell cell = row.getCell(colNum);

			// row.createCell((short) colNum).setCellValue(value);
			// Cell cell = row.getCell(colNum);

			// System.out.println("The Row Number " + lastRow +"is created" );

			// Cell cell;

			// cell = sheet.getRow(16).getCell(0);
			// System.out.println("Inside colnum =0 if condition");

			// cell.setCellFormula("IF(INDIRECT("+ "\""+ "V"+ "\""+
			// "&ROW()) = MAX($V:V),"+ "\""+ "Test1"+ "\" " +","+"\""+
			// "Test"+" \""+")");
			// evaluator.evaluateFormulaCell(cell);

			// System.out.println("Inside else condition");

			// /cell = sheet.getRow(i).getCell(2);
			// /cell.setCellValue(cell.getNumericCellValue() * 2);
			// /cell = sheet.getRow(2).getCell(2);
			// /cell.setCellValue(cell.getNumericCellValue() * 6);
			// /cell = sheet.getRow(3).getCell(2);
			// cell.setCellFormula("IF(INDIRECT("+ "\""+ "V"+ "\""+
			// "&ROW()) = MAX($V:V),"+ "\""+ "New"+ "\" " +","+"\""+
			// "Existing"+" \""+")");
			// evaluator.evaluateFormulaCell(cell);

			// cell.setCellValue(cell.getNumericCellValue() * 2);

			// System.out.println("The last Rownum is : "+
			// sheet.getLastRowNum());

			// FileOutputStream outFile =new FileOutputStream(new
			// File("/Users/droy2/Documents/ellipse/Test1.xlsx"));

			// workbook.write(outFile);
			FileOutputStream outFile = new FileOutputStream(new File(
					filename));
			workbook.write(outFile);
			outFile.close();
			file.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

	}

	private String NVL(String string, int i) {
		String result;
		
		if (string == null | string.length() == 0) {
			
			result = "0.00";
             
         } else { 
        	 result = string;
         }
		return result;
	}


	public void formulaEvaluator() throws IOException {

		FileInputStream fis = new FileInputStream("/somepath/test.xls");
		XSSFWorkbook wb = new XSSFWorkbook(fis); // or new
													// XSSFWorkbook("/somepath/test.xls")
		XSSFSheet sheet = wb.getSheetAt(0);
		FormulaEvaluator evaluator = wb.getCreationHelper()
				.createFormulaEvaluator();

		// suppose your formula is in B3
		CellReference cellReference = new CellReference("B3");
		Row row = sheet.getRow(cellReference.getRow());
		Cell cell = row.getCell(cellReference.getCol());

		if (cell != null) {
			switch (evaluator.evaluateInCell(cell).getCellType()) {
			case Cell.CELL_TYPE_BOOLEAN:
				System.out.println(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				System.out.println(cell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				System.out.println(cell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BLANK:
				break;
			case Cell.CELL_TYPE_ERROR:
				System.out.println(cell.getErrorCellValue());
				break;

			// CELL_TYPE_FORMULA will never occur
			case Cell.CELL_TYPE_FORMULA:
				break;
			}
		}
	}

	/*
	 * public static void main(String[] args) throws WriteException,
	 * IOException, BiffException {
	 * 
	 * UpdateExcl ue = new UpdateExcl(); ue.Update("Test", 24, 0);
	 * 
	 * }
	 */

}
