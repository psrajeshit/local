package test;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ClientsExcel {
	 @SuppressWarnings("null")
	public static void main(String[] args) {
	try 
	{ 
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");  
		
		 Connection con=DriverManager.getConnection("jdbc:sqlserver://127.0.0.1:1433;databaseName=FTS","sa","Idea@123");

		File excel = new File("C://Users/ISC10084/Desktop/Clients.xlsx");
			
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
	/*XSSFCell c= row.getCell((short) 4);
	System.out.println(c);
	//String c1 = c.getStringCellValue();
	XSSFCell d = row.getCell((short) 5);
	System.out.println(d);
	//String d1 = d.getStringCellValue();
	XSSFCell e = row.getCell((short) 6);
	System.out.println(e);
	//String e1 = e.getStringCellValue();
*/	XSSFCell f = row.getCell((short) 7);
	System.out.println(f);
	String f1 = f.getStringCellValue();
	XSSFCell g = row.getCell((short) 8);
	System.out.println(g);
	String g1 = g.getStringCellValue();
	XSSFCell h = row.getCell((short) 9);
	System.out.println(h);
	String h1 = h.getStringCellValue();
	XSSFCell i = row.getCell((short) 10);
	System.out.println(i);
	String i1 = i.getStringCellValue();
	XSSFCell j = row.getCell((short) 11);
	System.out.println(j);
	String j1 = j.getStringCellValue();
	XSSFCell k = row.getCell((short) 12);
	System.out.println(k);
	String k1 = k.getStringCellValue();
	XSSFCell l = row.getCell((short) 13);
	System.out.println(l);
	String l1 = l.getStringCellValue();
	XSSFCell m = row.getCell((short) 14);
	System.out.println(m);
	String m1 = m.getStringCellValue();
	XSSFCell n = row.getCell((short) 15);
	System.out.println(n);
	String n1 = n.getStringCellValue();
	XSSFCell o = row.getCell((short) 16);
	System.out.println(o);
	String o1 = o.getStringCellValue();
	XSSFCell p = row.getCell((short) 17);
	System.out.println(p);
	String p1 = p.getStringCellValue();
	XSSFCell q = row.getCell((short) 18);
	System.out.println(q);
	String q1 = q.getStringCellValue();
	XSSFCell r = row.getCell((short) 19);
	System.out.println(r);
	String r1 = r.getStringCellValue();
	XSSFCell v = row.getCell((short) 20);
	System.out.println(v);
	String v1 = v.getStringCellValue();
	XSSFCell t = row.getCell((short) 21);
	System.out.println(t);
	String t1 = t.getStringCellValue();
	XSSFCell u = row.getCell((short) 22);
	System.out.println(u);
	String u1 = u.getStringCellValue();

	// for row 0 end 1
	List<String> compny = new ArrayList<String>();
	int z =0;
	int x =0;
	for(int s=3;s<6;s++)// total no of row.
	{
	XSSFRow row1=sheet.getRow(s);
	  
	//Date
	//System.out.println(row.getRowNum());
	XSSFCell cella1 = row1.getCell((short) 0);
	if(cella1.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		cella1.setCellType(Cell.CELL_TYPE_STRING);
	}
	
	String Company="";
	String comp="";
	if(cella1!=null)
	{
		Company =cella1.getStringCellValue();
		
		System.out.println(Company);
		comp=Company;
	}
	
	//System.out.println(a2);
	String status ="";
	XSSFCell a21 = row1.getCell((short) 3);// status
	if(cella1.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		cella1.setCellType(Cell.CELL_TYPE_STRING);
	}
	System.out.println(a21);
	if(a21==null || (a21.getCellType() == Cell.CELL_TYPE_BLANK))
	{
	}
	else{
		String town=a21.getStringCellValue();
		if(town.equalsIgnoreCase("Yes")){
			status = "Contracted";
		} else {
			status = "Trial";
		}
		System.out.println(status);
	}

	PreparedStatement stmt;
	if(!compny.contains(Company)){
		x++;
		//Company Info
		stmt=con.prepareStatement("insert into  tbl_companyInfo values(?,?,?,?,?,?,?)"); 
		//1 specifies the first parameter in the query
		stmt.setTimestamp(1, new Timestamp(GregorianCalendar.getInstance().getTimeInMillis())); //creationDate
		stmt.setString(2,Company); //companyName 
		stmt.setString(3,"");  //address
		stmt.setString(4,"");  //email
		stmt.setString(5,"");  //phone
		stmt.setString(6,status);  //status
		stmt.setString(7,"");  //expiry
		stmt.executeUpdate(); 
		compny.add(Company);
	}
	
	
	
	//Client info
	XSSFCell c21 = row1.getCell((short) 1);//client name
	if(cella1.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		cella1.setCellType(Cell.CELL_TYPE_STRING);
	}
	String client="";
	if(c21!=null && !(c21.getCellType() == Cell.CELL_TYPE_BLANK))
	{
		client =c21.getStringCellValue();
		
		System.out.println(client);
	}
	 //Incident type

	
	String email="";
	XSSFCell cellb21 = row1.getCell((short) 2);// client email
	if(cella1.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		cella1.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(cellb21==null || (c21.getCellType() == Cell.CELL_TYPE_BLANK))
	{
	email="";
	}
	else{
		String town=cellb21.getStringCellValue();
		email=town;
		System.out.println(email);
	}
	z++;
	//Client
	stmt=con.prepareStatement("insert into  tbl_clients values(?,?,?,?)"); 
	stmt.setString(1,client); //Client name 
	stmt.setString(2,email);  //email
	stmt.setString(3,"");  //phone
	stmt.setInt(4,x);  //comapny
	stmt.executeUpdate(); 

	
	/*stmt=con.prepareStatement("insert into  FK_TBLUSER values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"); 
	stmt.setString(1,client); //Client name 
	stmt.setString(2,email);  //email
	stmt.setString(3,"");  //phone
	stmt.setInt(4,x);  //comapny
	stmt.executeUpdate(); */

	
	//Burundi
	XSSFCell d21= row1.getCell((short) 7);//tier1
	String tiers1="";
	String tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	XSSFCell d22= row1.getCell((short) 8);//tier2
	String tiers2="";
	String tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	XSSFCell d23= row1.getCell((short) 9);//Advisories
	String advis="";
	String advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	/*if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Burundi"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	
	//Cameroon
	d21= row1.getCell((short) 10);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 11);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 12);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Cameroon"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	/*//Comoros
	d21= row1.getCell((short) 13);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 14);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 15);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Comoros"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	//Djibouti
	/*d21= row1.getCell((short) 16);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 17);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 18);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Djibouti"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	//Democratic Republic of the Congo
	d21= row1.getCell((short) 19);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 20);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 21);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	XSSFCell d24= row1.getCell((short) 22);//SitRep
	String sitRep="";
	String weekly=null;
	if(d24.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d24.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d24==null || d24.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		sitRep = d24.getStringCellValue();
		if(sitRep.equalsIgnoreCase("Yes")){
			weekly = "Weekly";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")
			 || sitRep.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
		if(weekly != null){
			if(tier1 != null || tier2 != null || advisory != null){
				tier += ",";
			}
			tier += weekly;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Democratic Republic of the Congo"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	//Eritrea
	/*d21= row1.getCell((short) 23);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 24);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 25);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Eritrea"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	//Ethiopia
	d21= row1.getCell((short) 26);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 27);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 28);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Ethiopia"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	
	//Kenya
	d21= row1.getCell((short) 29);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 30);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption";
		}
		
	}
	
	d22= row1.getCell((short) 31);//tier3
	String tiers3="";
	String tier3=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers3 = d22.getStringCellValue();
		if(tiers3.equalsIgnoreCase("Yes")){
			tier3 = "Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 32);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	d24= row1.getCell((short) 33);//SitRep
	sitRep="";
	weekly=null;
	if(d24.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d24.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d24==null || d24.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		sitRep = d24.getStringCellValue();
		if(sitRep.equalsIgnoreCase("Yes")){
			weekly = "Weekly";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")
			 || sitRep.equalsIgnoreCase("Yes")|| tiers3.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(tier3 != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += tier3;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null || tier3 != null){
				tier += ",";
			}
			tier += advisory;
		}
		if(weekly != null){
			if(tier1 != null || tier2 != null || advisory != null || tier3 != null){
				tier += ",";
			}
			tier += weekly;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Kenya"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	//Madagascar
	/*d21= row1.getCell((short) 34);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 35);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 36);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Madagascar"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//Malawi
	d21= row1.getCell((short) 37);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 38);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 39);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Malawi"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	
	//Mozambique
	d21= row1.getCell((short) 40);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 41);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 42);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	d24= row1.getCell((short) 43);//SitRep
	sitRep="";
	weekly=null;
	if(d24.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d24.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d24==null || d24.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		sitRep = d24.getStringCellValue();
		if(sitRep.equalsIgnoreCase("Yes")){
			weekly = "Weekly";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")
			 || sitRep.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
		if(weekly != null){
			if(tier1 != null || tier2 != null || advisory != null){
				tier += ",";
			}
			tier += weekly;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Mozambique"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//Nigeria
	d23= row1.getCell((short) 44);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes")){
		String tier = "";
		if(advisory != null){
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Nigeria"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//Rwanda
	/*d21= row1.getCell((short) 45);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 46);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 47);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Rwanda"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//Seychelles
	d21= row1.getCell((short) 48);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 49);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 50);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Seychelles"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	
	//Somalia
	d21= row1.getCell((short) 51);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 52);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 53);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	d24= row1.getCell((short) 54);//SitRep
	sitRep="";
	weekly=null;
	if(d24.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d24.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d24==null || d24.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		sitRep = d24.getStringCellValue();
		if(sitRep.equalsIgnoreCase("Yes")){
			weekly = "Weekly";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")
			 || sitRep.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
		if(weekly != null){
			if(tier1 != null || tier2 != null || advisory != null){
				tier += ",";
			}
			tier += weekly;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Somalia"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//South Sudan
	/*d21= row1.getCell((short) 58);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 59);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 60);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"South Sudan"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//Sudan
	d21= row1.getCell((short) 61);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 62);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 63);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Sudan"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	
	//Tanzania
	d21= row1.getCell((short) 64);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 65);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 66);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	d24= row1.getCell((short) 67);//SitRep
	sitRep="";
	weekly=null;
	if(d24.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d24.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d24==null || d24.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		sitRep = d24.getStringCellValue();
		if(sitRep.equalsIgnoreCase("Yes")){
			weekly = "Weekly";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")
			 || sitRep.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
		if(weekly != null){
			if(tier1 != null || tier2 != null || advisory != null){
				tier += ",";
			}
			tier += weekly;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Tanzania"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//Uganda
	/*d21= row1.getCell((short) 68);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 69);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 70);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Uganda"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}

	
	//Zambia
	d21= row1.getCell((short) 74);//tier1
	tiers1="";
	tier1=null;
	if(d21.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d21.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d21==null || d21.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers1 = d21.getStringCellValue();
		if(tiers1.equalsIgnoreCase("Yes")){
			tier1 = "Catastrophic event of international significance. E.g Complex Attack";
		}
		
	}
	
	d22= row1.getCell((short) 75);//tier2
	tiers2="";
	tier2=null;
	if(d22.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d22.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d22==null || d22.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	System.out.println("");
	}
	else{
		tiers2 = d22.getStringCellValue();
		if(tiers2.equalsIgnoreCase("Yes")){
			tier2 = "Crime/ Terrorist activity and incidents of moderate disruption,Civil unrest/ RTAs/ Low level disruption";
		}
		
	}
	
	d23= row1.getCell((short) 76);//Advisories
	advis="";
	advisory=null;
	if(d23.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d23.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d23==null || d23.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		advis = d23.getStringCellValue();
		if(advis.equalsIgnoreCase("Yes")){
			advisory = "Advisory";
		}
	}
	
	d24= row1.getCell((short) 77);//SitRep
	sitRep="";
	weekly=null;
	if(d24.getCellType() == Cell.CELL_TYPE_NUMERIC) {
		d24.setCellType(Cell.CELL_TYPE_STRING);
	}
	if(d24==null || d24.getCellType() == Cell.CELL_TYPE_BLANK)
	{
	}
	else{
		sitRep = d24.getStringCellValue();
		if(sitRep.equalsIgnoreCase("Yes")){
			weekly = "Weekly";
		}
	}
	
	if(advis.equalsIgnoreCase("Yes") || tiers1.equalsIgnoreCase("Yes") || tiers2.equalsIgnoreCase("Yes")
			 || sitRep.equalsIgnoreCase("Yes")){
		String tier = "";
		if(tier1 != null){
			tier += tier1;
		}
		if(tier2 != null){
			if(tier1 != null){
				tier += ",";
			}
			tier += tier2;
		}
		if(advisory != null){
			if(tier1 != null || tier2 != null){
				tier += ",";
			}
			tier += advisory;
		}
		if(weekly != null){
			if(tier1 != null || tier2 != null || advisory != null){
				tier += ",";
			}
			tier += weekly;
		}
			
		//Client details
		stmt=con.prepareStatement("insert into  tbl_clientsDetails values(?,?,?,?,?)"); 
		stmt.setString(1,"Zambia"); //Country 
		stmt.setString(2,"");  //province
		stmt.setString(3,tier);  //tier
		stmt.setInt(4,z);  //client
		stmt.setString(5,"Email,Push Notifications");  //Notification
		stmt.executeUpdate(); 

	}*/

	
	
	
// Connection con=DriverManager.getConnection("jdbc:sqlserver://127.0.0.1:1433;databaseName=test_db","sa","ispl2016");

	
}
	System.out.println("records inserted row no:"+z);  
	con.close();
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

	