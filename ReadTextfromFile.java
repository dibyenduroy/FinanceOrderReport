package apachaePOI;

import java.io.*;

//import javaExcelWriter.WriteExcel;

public class ReadTextfromFile extends UpdateExcl {

	public static void main(String[] args) {

		ReadTextfromFile objupdate = new ReadTextfromFile();

		int lastRow;
		int yestardayLastRow;
		int lastRow0;
		int yestardayLastRow0;

		// args[0] should be the excel sheet file name : exp
		// "/Users/droy2/Documents/ellipse/Test2.xlsx"
		// args[1] //Exception Report Text File Name : exp
		// "/Users/droy2/Documents/ellipse/ExceptionReportAdditionsWithLines20130904.txt"
		// args[2] // New Siebel Orders Text File names : exp
		// "/Users/droy2/Documents/ellipse/NewSiebelOrders20130809.txt"

		try {
			//Added on Oct8th 2013 changed the Sheet Number from 4 to 2
			System.out.println("Last Row Yesterday is :"
					+ objupdate.getYesterdayLastRow(2, "N", 12345,args[0]));
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		try {
			lastRow = objupdate.getLastRownum(0, args[0]);// args [0] is will be
															// the Excel File
															// Name
			
			
			//Added on Oct8th 2013 changed the Sheet Number from 4 to 2
			yestardayLastRow = objupdate.getYesterdayLastRow(2, "N", lastRow,args[0]);
			
			objupdate.updateSheetRow(0, yestardayLastRow, lastRow, args[0],
					"Existing");
			//Added on Oct8th 2013 changed the Sheet Number from 4 to 2
			yestardayLastRow = objupdate.getYesterdayLastRow(2, "Y", lastRow,args[0]);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// ReadTextfromFile rd = new ReadTextfromFile();
		// rd.ReadFileandUpdate("/Users/droy2/Documents/ellipse/NewSiebelOrders20130809.txt",1);
		// rd.ReadFileandUpdate("/Users/droy2/Documents/ellipse/ExceptionReportAdditionsWithLines20130819.txt",0);

		// The name of the file to open.
		String fileName = args[1];
		ReadTextfromFile obj1 = new ReadTextfromFile();
		String fileName1 = args[2];

		// This will reference one line at a time
		String line = null;
		try {
			// FileReader reads text files in the default encoding.
			FileReader fileReader = new FileReader(fileName);
			FileReader fileReader1 = new FileReader(fileName1);
			// Always wrap FileReader in BufferedReader.
			BufferedReader bufferedReader = new BufferedReader(fileReader);
			BufferedReader bufferedReader1 = new BufferedReader(fileReader1);

			// Reading the New SiebelOrders File Line by Line

			while ((line = bufferedReader1.readLine()) != null) {

				String fields[] = line.split("\\t");

				int row = obj1.getLastRownum(1, args[0]) + 1;

				// System.out.println("The Field length is "+ fields.length);

				/*
				 * for (int i =0; i<fields.length;i++) {
				 * 
				 * String value1 = fields[i]; System.out.println("value1 is "+
				 * value1+"LastRow" +row+"The Column Number is "+i);
				 * 
				 * System.out.println(fields[i]);
				 * 
				 * obj1.Update(value1, row, i);
				 * 
				 * 
				 * }
				 */

				obj1.Update(1, row, fields.length, fields,args[0]);

			}

			// Reading the Exceptions File Line by Line
			while ((line = bufferedReader.readLine()) != null) {

				String fields[] = line.split("\\t");

				int row = obj1.getLastRownum(0, args[0]) + 1;

				// System.out.println("The Field length is "+ fields.length);

				/*
				 * for (int i =0; i<fields.length;i++) {
				 * 
				 * String value1 = fields[i]; System.out.println("value1 is "+
				 * value1+"LastRow" +row+"The Column Number is "+i);
				 * 
				 * System.out.println(fields[i]);
				 * 
				 * obj1.Update(value1, row, i);
				 * 
				 * 
				 * }
				 */

				obj1.Update(0, row, fields.length, fields,args[0]);

			}

			// Always close files.
			bufferedReader.close();
			bufferedReader1.close();
		}

		catch (IOException ex) {
			System.out.println("Error reading file '" + fileName + "'");
			// Or we could just do this:
			// ex.printStackTrace();
		}

		// Get Siebel rows
		try {

			lastRow0 = objupdate.getLastRownum(0, args[0]);// last row from
															// sheet0
			//Updated on 8th Oct 2013 changed Row number from 4 to 2
			yestardayLastRow0 = objupdate.getYesterdayLastRow(2, "N", lastRow0,args[0]);// yesterday
																				// last
																				// row
																				// from
																				// sheet
																				// 1

			lastRow = objupdate.getLastRownum(1, args[0]);// last row from
															// sheet1
			yestardayLastRow = objupdate.getSiebelYesterdayLastRow(2, "N",
					lastRow,args[0]);// yesterday last row from sheet

			System.out.println("The last first row of sheet0 " + lastRow0);
			System.out.println("The last first row of sheet1 " + lastRow);
			//yestardayLastRow - 6000
			//yestardayLastRow - 1
			System.out.println("yestardayLastRow is  "+ yestardayLastRow);
			System.out.println("lastRow is  "+ lastRow);
			//lastRow0
			//Changed on Oct8th 2013 replaced 3279 by 1893
			//Changed on Jan6th 2014 replaced 1893 by 3
			objupdate.updateSiebelFlag(0, 2, lastRow0,
					yestardayLastRow - 250, lastRow, args[0], "Existing");
			//Changed on Oct8th 2013 replaced 4 by 2
			yestardayLastRow = objupdate.getSiebelYesterdayLastRow(2, "Y",
					lastRow,args[0]);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}