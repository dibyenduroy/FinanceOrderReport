package javaExcelWriter;

import java.awt.Label;  
import java.io.File;  
import java.io.IOException;

import jxl.Cell;   
import jxl.CellType;  
import jxl.LabelCell;  
import jxl.NumberCell;  
import jxl.Sheet;  
import jxl.Workbook;  
import jxl.read.biff.BiffException;  
import jxl.write.WritableCell;  
import jxl.write.WritableSheet;  
import jxl.write.WritableWorkbook;  
import jxl.write.WriteException;  
import jxl.write.biff.RowsExceededException;  

public class ExcelJxl {

/**
 * @param args
 * @throws IOException 
 * @throws BiffException 
 * @throws WriteException 
 * @throws RowsExceededException 
 */
public static void main(String[] args) throws  IOException, RowsExceededException, WriteException, BiffException {
    // TODO Auto-generated method stub
         ExcelJxl.WriteFile("/Users/droy2/Documents/workspace/lars.xlsx");
}

public static void WriteFile(String path) throws BiffException, IOException, RowsExceededException, WriteException{

Workbook wb=Workbook.getWorkbook(new File(path));

WritableWorkbook copy=Workbook.createWorkbook(new File("/Users/droy2/Documents/workspace/larscopy.xlsx"),wb);
WritableSheet sheet = copy.getSheet(1); 
WritableCell cell = sheet.getWritableCell(0,0); 
String S="nimit";
if (cell.getType() == CellType.LABEL) 
{ 
  LabelCell l = (LabelCell) cell; 
  //l.setString(S); 
}
copy.write(); 
copy.close();
wb.close();

}}