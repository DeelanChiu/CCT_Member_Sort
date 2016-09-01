//for excel reading
import java.util.Date;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
//for URL access
import java.net.*;
import java.io.*;
import java.lang.*;

public class ListservSplit {
	public static void main (String args[]) throws BiffException, IOException, RowsExceededException, WriteException, Exception{
		File member_manifest = new File("D:/Program Files/eclipse/members_export.xls");
		Workbook workbook = Workbook.getWorkbook(member_manifest);
		Sheet sht = workbook.getSheet(0);
		String netID = null;
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
				email=sht.getCell(0,memberCount++);	
				//want to get rid of @cornell.edu...  12 char
				//so truncate to string.length - 12
				//next: find index of @ and cut off everything after
				netID = email.getContents().substring(0,email.getContents().length()-12);
				//System.out.println(netID);
				
				//if it's not a cornell account, it'll try to use name for search instead
				URL web = new URL("https://www.cornell.edu/search/people.cfm?netid="+netID);
				//using name:
				//https://www.cornell.edu/search/people.cfm?q=Dylan+Chiu
				//https://www.cornell.edu/search/people.cfm?q='first'+'last'
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
		        int colCount = 0;
		        Cell rowCell = sht.getCell(colCount,memberCount-1);
		        Label cell;
		        try{
		        	while (!rowCell.getContents().isEmpty()){
		        		
		        		if (student && !alum){
		        			Ssheet.addCell(new Label(colCount, studCount, rowCell.getContents()));
		        			System.out.println("printed "+rowCell.getContents()+" in student");
		        		} else if (alum && !student){
		        			Asheet.addCell(new Label(colCount, alumCount, rowCell.getContents()));
		        			System.out.println("printed  "+rowCell.getContents()+" in alum");
		        		}
		        		rowCell=sht.getCell(colCount++, memberCount-1);
		        	}
		        } catch (ArrayIndexOutOfBoundsException e) {}
		        
		        alum = false;
		        student = false;
		        
		        read.close();
				
				
			}
		} catch (ArrayIndexOutOfBoundsException e) {
			memberCount-=2;
			//System.out.println(memberCount);
			if (memberCount==0) {
				//System.out.println("Empty sheet");
				return;
			}
			studCount++;
			alumCount++;
		}
		
		
		alumL.write();
		studL.write();
		alumL.close();
		studL.close();
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
