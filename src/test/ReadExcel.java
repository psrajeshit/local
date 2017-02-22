package test;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ReadExcel {
	 public static void main(String[] args) {
	try 
	{ 
		File excel = new File("C://Users/ISC10084/Desktop/Tansania.xlsx");
			
	//File excel = new File("D://temp/ab6.xlsx"); //dataofsf
	FileInputStream fis = new FileInputStream(excel);
	XSSFWorkbook book = new XSSFWorkbook(fis);
	XSSFSheet sheet = book.getSheetAt(0);
	// for row 0
	XSSFRow row=sheet.getRow(0);
	
	//System.out.println(row.getRowNum());
	XSSFCell cellA1 = row.getCell((short) 0);
	System.out.println(cellA1);
//	String a1Val = cellA1.getStringCellValue();
	XSSFCell cellb1 = row.getCell((short) 1);
	System.out.println(cellb1);
	//String b1Vbl = cellb1.getStringCellValue();
	XSSFCell a = row.getCell((short) 2);
	System.out.println(a);
	//String a1 =a.getStringCellValue();
	XSSFCell b = row.getCell((short) 3);
	System.out.println(b);
	//String b1 = b.getStringCellValue();
	XSSFCell c= row.getCell((short) 4);
	System.out.println(c);
	//String c1 = c.getStringCellValue();
	XSSFCell d = row.getCell((short) 5);
	System.out.println(d);
	//String d1 = d.getStringCellValue();
	XSSFCell e = row.getCell((short) 6);
	System.out.println(e);
	//String e1 = e.getStringCellValue();
	XSSFCell f = row.getCell((short) 7);
	System.out.println(f);
	//String f1 = f.getStringCellValue();
	XSSFCell g = row.getCell((short) 8);
	System.out.println(g);
	//String g1 = g.getStringCellValue();
	XSSFCell h = row.getCell((short) 9);
	System.out.println(h);
	//String h1 = h.getStringCellValue();
	XSSFCell i = row.getCell((short) 10);
	System.out.println(i);
//	String i1 = i.getStringCellValue();
	XSSFCell j = row.getCell((short) 11);
	System.out.println(j);
	//String j1 = j.getStringCellValue();
	XSSFCell k = row.getCell((short) 12);
	System.out.println(k);
	//String k1 = k.getStringCellValue();

	// for row 0 end 1
	for(int s=1;s<1000;s++)// total no of row.
	{
	XSSFRow row1=sheet.getRow(s);
	  
	//source
	//System.out.println(row.getRowNum());
	XSSFCell cella1 = row1.getCell((short) 0);
	
	String source="";
	//System.out.println(cella1);
	String a2 = cella1.getStringCellValue();
	if(a2.equals("A"))
	{
		System.out.println("Live: Corroborated");
		source="Live: Corroborated";
	}
	if(a2.equals("B"))
	{
		System.out.println("Live: Uncorroborated");
		 source="Live: Uncorroborated";
	}
	if(a2.equals("C"))
	{
		System.out.println("Backlog");
		 source="Backlog";
	}
	
	//System.out.println(a2);
	String tow="";
	XSSFCell cellb21 = row1.getCell((short) 1);// town or city
	if(cellb21==null)
	{
	System.out.println("");
	tow="";
	}
	else{
		String town=cellb21.getStringCellValue();
		tow=town;
		System.out.println(town);
	}
	
	XSSFCell a21 = row1.getCell((short) 2);// country
	System.out.println(a21);
//	String c2 =a21.getStringCellValue();
	//System.out.println(c2);
	XSSFCell c21 = row1.getCell((short) 3);//Incident Type
	String Incident="";
	String inc="";
	if(c21!=null)
	{
		Incident =c21.getStringCellValue();
		if(Incident.equals("Robbery/Criminal Activity"))
		{
		System.out.println("Criminal Activity");
		inc="Criminal Activity";
		}
		else if(Incident.equals("Asset Attack"))
		{
		System.out.println("Criminal Activity");
		inc="Criminal Activity";
		}
		else if(Incident.equals("Targeted Killing"))
		{
		System.out.println("Assassination");
		inc="Assassination";
		}
		else if(Incident.equals("Demonstration"))
		{
		System.out.println("Protest");
		inc="Protest";
		}
		else
		{
		System.out.println(Incident);
		inc=Incident;
		}
	}
	 //Incident type

	//String d2 =c21.getStringCellValue();
	//System.out.println(d2);
	XSSFCell d21= row1.getCell((short) 4);//Impact
	String imp="";
	int tierwgt=0;
	 String Impact="";
	
	 if(d21!=null)
	 { Impact =d21.getStringCellValue();
	  if(Impact.equals("Low"))
		{
		System.out.println("Civil unrest/ RTAs/ Low level disruption");
		tierwgt=2;
		imp="Civil unrest/ RTAs/ Low level disruption";
		}
	 if(Impact.equals("Very Low"))
		{
		System.out.println("Civil unrest/ RTAs/ Low level disruption");
		tierwgt=2;
		imp="Civil unrest/ RTAs/ Low level disruption";
		}
	if(Impact.equals("Unknown"))
		{
		System.out.println("Civil unrest/ RTAs/ Low level disruption");
		tierwgt=2;
		imp="Civil unrest/ RTAs/ Low level disruption";
		}
	if(Impact.equals("Medium"))
	{
	System.out.println("Crime/terrorist activity and incidents of moderate disruption");
	tierwgt=3;
	imp="Crime/terrorist activity and incidents of moderate disruption";
	}
	if(Impact.equals("High"))
	{
	System.out.println("Catastrophic event of international significance. E.g Complex Attack");
	imp="Catastrophic event of international significance. E.g Complex Attack";
	tierwgt=5;
	}
	if(Impact.equals("Very High"))
	{
	System.out.println("Catastrophic event of international significance. E.g Complex Attack");
	imp="Catastrophic event of international significance. E.g Complex Attack";
	tierwgt=5;
	}}
	// for killed
	XSSFCell f21= row1.getCell((short) 5);
	String kil="";
	Integer killed=(int)f21.getNumericCellValue();
	if(killed<=10 )
	{
		 if(killed==-1)
		{
			System.out.println("UnKnown");
			kil="UnKnown";
		}
		 else
		 {
	
	kil=killed.toString();
	System.out.println(kil);
	}
	}
	
	else
	{
		System.out.println(">10");
		kil=">10";
	}
	
	// for injured
	XSSFCell g21 = row1.getCell((short) 6);// injured
	String inj="";
	Integer injured =(int)g21.getNumericCellValue();
	if(injured<=10 )
	{
		if(injured==-1)
		{
			System.out.println("UnKnown");
			inj="UnKnown";
		}
		else
		{
			inj=injured.toString();
	System.out.println(injured);
	}
	}
	 
	else
	{
		System.out.println(">10");
		inj=">10";
	}

	XSSFCell h21= row1.getCell((short) 7);
	System.out.println("hello"+h21);
	String opp="";
	 //opposition Insurgents
	if(h21==null)
	{
	System.out.println("None");
	opp="None";
	}
	else
	{
		String opposition=h21.getStringCellValue();
		if(opposition.equals("Less than 10"))
		{
			opp="<10";
		System.out.println("<10");
		}
		 if(opposition.equals("Less than 50"))
		{
			 opp=">25";
			System.out.println(">25");
		}
		 if(opposition.equals("Greater than 50"))
		{
			 opp=">25";
			System.out.println(">25");
		}
		 if(opposition.equals("Unknown at this time"))
		{
			 opp="Unknown";
			System.out.println("Unknown");
		}
	}
// for accuracy
	XSSFCell i21= row1.getCell((short) 8);//
	String acc="";
	String accuracy="";
	if(i21!=null)
	{
		accuracy=i21.getStringCellValue();
	if(accuracy.equals("accuracy_high"))
	{
		acc="High - plotted by known area or neighbourhood of a city/town";
		System.out.println("High - plotted by known area or neighbourhood of a city/town");
	}
	if(accuracy.equals("accuracy_medium"))
	{
		System.out.println("Medium - plotted by city/town");
		acc="Medium - plotted by city/town";
	}
	if(accuracy.equals("accuracy_very_high"))
	{
		System.out.println("Very High - plotted by known building or landmark");
		acc="Very High - plotted by known building or landmark";
	}
	if(accuracy.equals("accuracy_low"))
	{
		System.out.println("Low - plotted by province or region");
		acc="Low - plotted by province or region";
	}}
	
// for date
	XSSFCell j21= row1.getCell((short) 9);
	String ans=j21.getStringCellValue();
	/*Integer date=ans.getDate();
	String dat=date.toString();
	  Integer month=ans.getMonth()+1;
	  String canmonth="";
	  if(month<10)
	  {
		  
		  canmonth="0"+month;
	  }
	  else
	  {
		  canmonth=month.toString(); 
	  }
	  String candate="";
	  if(date<10)
	  {

      candate="0"+date;
	  }
	  else
	  {
		  candate=date.toString();
	  }
	
	  Integer yrs=ans.getYear();
	 String yr=yrs.toString();
	char ch1=yr.charAt(1);
	char ch2=yr.charAt(2);
	 System.out.println(candate+"/"+canmonth+"/20"+ch1+ch2);*/
	// String dmyr=candate+"/"+canmonth+"/20"+ch1+ch2;
	 String dmyr=ans;

	XSSFCell k21= row1.getCell((short) 10);//lat
   Double lat=k21.getNumericCellValue();
   String latitude=lat.toString();
	System.out.println(lat);//
	XSSFCell l21= row1.getCell((short) 11);//log
	Double log=l21.getNumericCellValue();
	String longitude=log.toString();
	//System.out.println(log);
	XSSFCell m21= row1.getCell((short) 12);
	System.out.println(m21);
	XSSFCell sf= row1.getCell((short) 13);
	String sfdata="";
	if(sf==null)
	{
		sfdata="";
	}
	else
	{
	sfdata=sf.getStringCellValue(); 
	}
	System.out.println(sf);
	// incident weight
	String[] myStringayArr1 = new String[]{"Air Strike","Complex Attack","IED","Kidnapping","Maritime - Hijack","Shelling","Murder"};//5
	String[] myStringayArr2 = new String[]{"Riot","Livestock Rustling","Shooting","Assassination","Robbery","Threat Warning","Smuggling","Poaching","Natural Disaster","Mob Violence","Maritime - Illegal Boarding","Maritime - Attack","Targeted Killing","Government Security Response"};//4
	String[] myStringayArr3 = new String[]{"Maritime - Robbery","Other","Protest","Route Obstruction","SGBV","Criminal Activity","Domestic Violence","Trafficking","Robbery/Criminal Activity","Asset Attack","Demonstration"};//3
	String[] myStringayArr4 = new String[]{"Maritime - Suspicious Activity/Approach","RTA"};//2

	int incidweight=0;
if(Incident.matches("Air Strike|Complex Attack|IED|Kidnapping|Maritime - Hijack|Shelling|Murder"))
{
	
	incidweight=5;
	}
if(Incident.matches("Riot|Livestock Rustling|Shooting|Assassination|Robbery|Threat Warning|Smuggling|Poaching|Natural Disaster|Mob Violence|Maritime - Illegal Boarding|Maritime - Attack|Targeted Killing|Government Security Response"))
{

	incidweight=4;
	}
if(Incident.matches("Maritime - Robbery|Other|Protest|Route Obstruction|SGBV|Criminal Activity|Domestic Violence|Trafficking|Robbery/Criminal Activity|Asset Attack|Demonstration"))
{
	
	incidweight=3;
	}
if(Incident.matches("Maritime - Suspicious Activity/Approach|RTA"))
{
	incidweight=2;
	
	
}

	int kilweight = 0;
	if(killed==-1 || killed==0)
	{
		System.out.println("1");
		kilweight=1;
	}
	if(killed==1)
	{
		System.out.println("2");
		kilweight=2;
	}
	if(killed==2)
	{
		System.out.println("3");
		kilweight=3;
	}

	if(killed>=3 && killed<=8)
	{
		System.out.println("4");
		kilweight=4;
	}
	if(killed>=9)
	{
		System.out.println("5");
		kilweight=5;
	}
	// for injured
	int injweight = 0;
	if(injured==-1)
	{
		System.out.println("1");
		injweight=1;
	}
	if(injured==0 || injured==1)
	{
		
		System.out.println("1");
		injweight=1;
	}
	if(injured==2 || injured==3)
	{
		
		System.out.println("2");
		injweight=2;
	}
	if(injured==4 || injured==5)
	{
		
		System.out.println("3");
		injweight=3;
	}
	if(injured>=6 && injured<=9)
	{
		
		System.out.println("4");
		injweight=4;
	}
	if(injured>=10)
	{
		System.out.println("5");
		injweight=5;
	}
	
	



Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");  
	
Connection con=DriverManager.getConnection("jdbc:sqlserver://62.12.115.183:1433;databaseName=BREATHE","sa","Breathe@DB");
 //Connection con=DriverManager.getConnection("jdbc:sqlserver://127.0.0.1:1433;databaseName=FTS","sa","Idea@123");

PreparedStatement stmt=con.prepareStatement("insert into  incident_list values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); 
	//1 specifies the first parameter in the query
	stmt.setTimestamp(1, new Timestamp(GregorianCalendar.getInstance().getTimeInMillis())); //creationDate
	stmt.setString(2,""); //incidentId 
	stmt.setString(3,"DataPort");  //createdUser
	stmt.setString(4,"DataPort");  //currentUser
	stmt.setString(5,"Tanzania");  //country
	stmt.setString(6,tow);  //district
	stmt.setString(7,"");  //town
	stmt.setString(8,acc);  //accuracy
	stmt.setString(9,"");  //intendedTarget
	stmt.setString(10,opp);  //opposition
	stmt.setString(11,source);  //source
	stmt.setString(12,inc); //incidentType 
	stmt.setString(13,kil);  //killed
	stmt.setString(14,inj);//injured
	stmt.setString(15,sfdata);//incidentDetails
	stmt.setString(16,"Approved");//status
	stmt.setString(17,latitude);//latitude
	stmt.setString(18,longitude);//longitude
	stmt.setString(19,imp);//tiers
	stmt.setString(20,"");//sfCommentary
	stmt.setString(21,"");//keywords
	stmt.setString(22,dmyr);//incidentDate
	stmt.setString(23,"");//incidentTime
	stmt.setString(24,"");//approximateTime
	stmt.setString(25,"");//reviewComment
	stmt.setBoolean(26,false);//images
	stmt.setString(27,"");//mapUrl
	stmt.setString(28,"");//clientStatus
	stmt.setString(29,"");//clientComment
	stmt.setInt(30,incidweight);//incidentWeight
	stmt.setInt(31,tierwgt);//tierWeight
	stmt.setInt(32,kilweight);//killedWeight
	stmt.setInt(33,injweight);//injuredWeig
	stmt.setBoolean(34,true);//editIncident
	 
	int ins=stmt.executeUpdate(); 
	System.out.println("records inserted row no:"+s);  
con.close();
}
	 
//	Iterator<Row> itr = sheet.iterator(); // Iterating over Excel file in Java
	//while (itr.hasNext()) {
		
	//	Row row = itr.next(); // Iterating over each column of Excel file 
	//Iterator<Cell> cellIterator = row.cellIterator();
	//while (cellIterator.hasNext())
	//{ 
//	/	
	//Cell cell = cellIterator.next();
//	switch (cell.getCellType()) 
//	{
//	case 
//	Cell.CELL_TYPE_STRING: System.out.print(cell.getStringCellValue() + "\t");
//	break;
//	case
//	Cell.CELL_TYPE_NUMERIC: System.out.print(cell.getNumericCellValue() + "\t");
	//break; 
	//case Cell.CELL_TYPE_BOOLEAN: System.out.print(cell.getBooleanCellValue() + "\t"); 
//	break; 
	
	//default:
//		}
//	} 
//	System.out.println("");
	//}

	}
	catch(Exception e)
	{
		System.out.println("Stop inserting......");
	}
	/*try{  
		System.out.println("start");
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");  
		
		Connection con=DriverManager.getConnection("jdbc:sqlserver://127.0.0.1:1433;databaseName=test_db","sa","ispl2016");
		  
		System.out.println("start connection");
		PreparedStatement stmt=con.prepareStatement("insert into  bike values(?,?,?)");  
		stmt.setInt(1,4);//1 specifies the first parameter in the query
		stmt.setInt(2,1234);
		stmt.setString(3,"Ratan");  
		  
		int i=stmt.executeUpdate();  
		System.out.println(i+" records inserted");  
		  
		con.close();  
		  
		}catch(Exception e){ System.out.println(e);}  */
		  
		 
	
	}

	private static java.sql.Date getDate() {
		// TODO Auto-generated method stub
		return null;
	}}

	