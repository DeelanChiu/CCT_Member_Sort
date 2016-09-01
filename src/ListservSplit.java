//for excel reading
import java.util.Date;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
//for URL access
import java.net.*;
import java.io.*;
import java.lang.*;
//for writing to files
import java.io.FileWriter;
import java.io.PrintWriter;
import java.io.IOException;


public class ListservSplit {
	public static void main (String args[]) throws BiffException, IOException, RowsExceededException, WriteException, Exception{
		File member_manifest = new File("members_export.xls");
		Workbook workbook = null;
		try {
			 workbook = Workbook.getWorkbook(member_manifest);
		} catch (FileNotFoundException e){
			JFrame frame = new JFrame("Can't find file");
			JOptionPane.showMessageDialog(frame, "A member manifest can't be found in the"
					+ " directory of the program!");
			System.exit(0);
		}
		
		PrintWriter printLine = new PrintWriter( new FileWriter("report", false));
		
		
		Sheet sht = workbook.getSheet(0);
		int memberCount = 1;
		int studCount = 0;
		int alumCount = 0;
		Cell email = sht.getCell(0,memberCount);
		
		WritableWorkbook alumL = Workbook.createWorkbook(new File("alumni_list.xls"));
		WritableSheet Asheet = alumL.createSheet("alumni list", 0);
		WritableWorkbook studL = Workbook.createWorkbook(new File("student_list.xls"));
		WritableSheet Ssheet = studL.createSheet("student list", 0);
		
		try {
			while (!email.getContents().isEmpty()){
				URL web;
				
				email=sht.getCell(0,memberCount++);	
				//find index of @ and cut off everything after
				int atInd = email.getContents().indexOf('@');
				//check if it's a cornell address
				if (email.getContents().substring(atInd, email.getContents().length()).contains("@cornell.edu")){
					//if it's a cornell address, it'll use netID to search up
					web = new URL("https://www.cornell.edu/search/people.cfm?netid="+
							email.getContents().substring(0,atInd));
				} else {
					//if it's not a cornell account, it'll try to use name for search instead
					//using name:
					//https://www.cornell.edu/search/people.cfm?q=Dylan+Chiu
					//https://www.cornell.edu/search/people.cfm?q='first'+'last'
					web = new URL("https://www.cornell.edu/search/people.cfm?q="+
							sht.getCell(1,memberCount).getContents().replace(' ', '+')
							+ sht.getCell(2,memberCount).getContents().replace(' ', '+'));
				}
				//netID = email.getContents().substring(0,atInd);
				//System.out.println(netID);
				

				BufferedReader read = new BufferedReader(
				        new InputStreamReader(web.openStream()));
				
				String inputLine;
				boolean alum = false;
				boolean student = false;
		        while ((inputLine = read.readLine()) != null && !student && !alum){
		            if (inputLine.contains("<td>student</td>")){
		            	student = true;
		            	studCount++;
		            	//System.out.println("student");
		            } else if (inputLine.contains("<td>alumni</td>")){
		            	alum = true;
		            	alumCount++;
		            	//System.out.println("alum");		            	
		            }
		            	
		        }
		        
		        if (!student && !alum){ 
		        	System.out.println("can't find: "+		        
		        		email.getContents().substring(0,atInd));
		        	printLine.println("can't find: "+		        
			        		email.getContents().substring(0,atInd));
		        } else {
			        int colCount = 0;
			        Cell rowCell = sht.getCell(colCount,memberCount-1);
			        Label cell;
			        try{
			        	while (!rowCell.getContents().isEmpty()){
			        		rowCell=sht.getCell(colCount, memberCount-1);
			        		if (student && !alum){
			        			Ssheet.addCell(new Label(colCount, studCount, rowCell.getContents()));
			        			//System.out.println("printed "+rowCell.getContents()+" in student with colCount "+colCount);
			        		} else if (alum && !student){
			        			Asheet.addCell(new Label(colCount, alumCount, rowCell.getContents()));
			        			//System.out.println("printed  "+rowCell.getContents()+" in alum");
			        		}
			        		colCount++;
			        	}
			        } catch (ArrayIndexOutOfBoundsException e) {}
			        
			        alum = false;
			        student = false;
		        }
		        
		        read.close();
				
				
			}
		} catch (ArrayIndexOutOfBoundsException e) {
			memberCount-=2;
			//System.out.println(memberCount);
			printLine.println(memberCount+" members");
			printLine.println(alumCount+" alumnis, "+studCount+" students");
			if (memberCount==0) {
				//System.out.println("Empty sheet");
				System.exit(0);;
			}
			studCount++;
			alumCount++;
		}
		
		JFrame frame = new JFrame("Done!");
		JOptionPane.showMessageDialog(frame, "Listserv processed! Please see the "
				+ " report for details on unsorted names.");
		
		
		alumL.write();
		studL.write();
		alumL.close();
		studL.close();
		printLine.close();
		/*
		WritableWorkbook writewkbk = Workbook.createWorkbook(new File("output.xls"));
		WritableSheet sheet = writewkbk.createSheet("First Sheet", 0);
		Label label = new Label(1, 1, "A label record"); 
		sheet.addCell(label);
		writewkbk.write();
		writewkbk.close();
		*/
	}
}
