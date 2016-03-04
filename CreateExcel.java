package javaExcelWriter;

import jxl.*;  
import jxl.read.biff.*;  

import java.io.File;  
import java.io.IOException;  

import jxl.write.WriteException;  

import java.util.Locale;  

import jxl.write.WritableSheet;  
import jxl.write.WritableWorkbook;  
import jxl.write.Label;   
   
public class CreateExcel {  
 /** 
  * @param args 
  */  
 public static void main(String[] args) throws IOException, WriteException, BiffException{  
  // TODO Auto-generated method stub  
  
final File  outputWorkbook;  
        try  
        {  
          
         File f1=new File("/Users/droy2/Documents/workspace/lars.xlsx");                                               //the excel sheet which contains data  
         WorkbookSettings ws=new WorkbookSettings();  
         ws.setLocale(new Locale("er","ER"));  
         Workbook workbook=Workbook.getWorkbook(f1,ws);   
           
         Sheet readsheet=workbook.getSheet(0);  
         System.out.println(readsheet.getName());    
         outputWorkbook = new File("/Users/droy2/Documents/workspace/lars2.xlsx");                            // the excel sheet where data is to copied  
         WritableWorkbook workbook1=Workbook.createWorkbook(outputWorkbook,workbook);  
          WritableSheet sheet1 = workbook1.getSheet("mysheet");  
         int i=0,j=0;                                                                                               //following code copies the data from sandy1 to sandy2.xls  
  
         for(j=0;j<(readsheet.getColumns());j++){          
         for(i=0;i<readsheet.getRows();i++)  {            
             
          if(readsheet.getCell(j,i).getType()==(CellType.LABEL))      
          {  
             
           Label Name=new Label(j,i,readsheet.getCell(j,i).getContents());  
            sheet1.addCell(Name);  
              
               System.out.println(readsheet.getCell(j,i).getContents());   
          }  
          if(readsheet.getCell(j,i).getType()==(CellType.NUMBER)) {     
                   
              String s = readsheet.getCell(j,i).getContents();        
         Label Name1=new Label(j,i,s);  
         sheet1.addCell(Name1);  
          System.out.println(readsheet.getCell(j,i).getContents());  
         }                 
            }  
         }  
         workbook1.write();  
      workbook1.close();  
           
        
        }  
        catch(IOException e)  
        {  
         e.printStackTrace();  
        }  
          
        catch(BiffException e)  
        {  
         e.printStackTrace();  
        }  
}  
}  