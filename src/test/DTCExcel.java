package test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.net.URLEncoder;
import java.nio.charset.Charset;
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



public class DTCExcel {
	 public static void main(String[] args) {
	try 
	{ 
		File excel = new File("C://Users/ISC10084/Desktop/DTCData.xlsx");
		String name = "C://Users/ISC10084/Desktop/DTCData.xlsx";
	if(name.length() >37){
		System.out.println(true);
	}else{
		System.out.println(false);
	}
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
	for(int s=1;s<500;s++)// total no of row.
	{
	XSSFRow row1=sheet.getRow(s);
	  
	//source
	//System.out.println(row.getRowNum());
	XSSFCell cella1 = row1.getCell((short) 0);//MemberId
	
	String menberId="";
	//System.out.println(cella1);
	String a2 = cella1.getStringCellValue();
	menberId = a2;
	
	//System.out.println(a2);
	String firstName="";
	XSSFCell cellb21 = row1.getCell((short) 1);// firstName
	String town=cellb21.getStringCellValue();
	firstName=town;
	System.out.println(firstName);
	
	XSSFCell a21 = row1.getCell((short) 2);// lastName
	System.out.println(a21);
	String lastName =a21.getStringCellValue();
	System.out.println(lastName);
	XSSFCell c21 = row1.getCell((short) 3);//Address1
	String address1="";
	if(c21!=null)
	{
		address1 =c21.getStringCellValue();
	}

	XSSFCell d21= row1.getCell((short) 4);//city
	String city="";
	
	if(d21!=null) { 
		city =d21.getStringCellValue();
	}

	XSSFCell f21= row1.getCell((short) 5); //state
	String state = f21.getStringCellValue();
	System.out.println(state);
	
	XSSFCell g21 = row1.getCell((short) 6);// zipCode
	String zipCode = "";
	Integer zip =(int)g21.getNumericCellValue();
	zipCode = zip.toString();
	if(zipCode.length() !=5 && zipCode.length()!=0){
		zipCode = "0"+zip.toString();
	}
	System.out.println(zipCode);

	XSSFCell h21= row1.getCell((short) 7); // telephone
	Integer tel =(int)h21.getNumericCellValue();
	String telephone=tel.toString();
	System.out.println(telephone);

	XSSFCell i21= row1.getCell((short) 8);// dob
	SimpleDateFormat sdf1 = new SimpleDateFormat("MM/dd/yyyy");
	Date date=i21.getDateCellValue();
	String dob = sdf1.format(date);
	System.out.println(dob);

	XSSFCell j21= row1.getCell((short) 9);// age
	Integer ag =(int)j21.getNumericCellValue();
	String age=ag.toString();

	XSSFCell k21= row1.getCell((short) 10);//email
    String email=k21.getStringCellValue();
    System.out.println(email);

	XSSFCell l21= row1.getCell((short) 11);//Infertility
	String infertility=l21.getStringCellValue();
    System.out.println(infertility);

    XSSFCell m21= row1.getCell((short) 12); //knownWin
    String knownWin=m21.getStringCellValue();
    System.out.println(knownWin);

	XSSFCell sf= row1.getCell((short) 13); // total
	Integer tot =(int)sf.getNumericCellValue();
	String total=tot.toString(); 
	System.out.println(total);
	
	/*String data = "https://pi.pardot.com/api/prospect/version/4/do/create/email/";
	data += email;
	data += "?first_name="+firstName.replaceAll(" ", "%20");
	data += "&last_name="+lastName.replaceAll(" ", "%20");
	data += "&source=ATS";
	data += "&campaign_id=692";
	data += "&address_one="+address1.replaceAll(" ", "%20");
	data += "&city="+city.replaceAll(" ", "%20");
	data += "&state="+state.replaceAll(" ", "%20");
	data += "&zip="+zipCode;
	data += "&phone="+telephone;
	data += "&Birthdate="+dob;
	data += "&Age="+age;
	data += "&Treated_for_Infertility_Score=Yes";
	data += "&score="+total;
	data += "&How_did_you_hear_about_WIN="+knownWin.replaceAll(" ", "%20");
	data += "&api_key=167562362d23c6d60e064d4b86aedc5b&user_key=074fba8df5a2b4d2ab9f244a8a314861";
	//String encoded = URLEncoder.encode(data, "UTF-8");
    URL url = new URL(data);

	 URLConnection conn = url.openConnection();
     conn.setDoOutput(true);
     //OutputStreamWriter wr = new OutputStreamWriter(conn.getOutputStream());
     //wr.write("");
    // wr.flush();

     // Get the response
     BufferedReader rd = new BufferedReader(new InputStreamReader(((HttpURLConnection) (new URL(data)).openConnection()).getInputStream(), Charset.forName("UTF-8")));

     //BufferedReader rd = new BufferedReader(new InputStreamReader(conn.getInputStream()));
     String line;
     String id = "";
     int z=1;
     while ((line = rd.readLine()) != null) {
         if(z == 4){
        	 line = line.replaceAll("<id>", "");
        	 line = line.replaceAll("</id>", "");
        	 line = line.replaceAll(" ", "");
        	 id = line;
        	 System.out.println("id is: "+id);
         }
         z++;
    	 System.out.println(line);
    	 
     }
     //wr.close();
     rd.close();
     String datas = "https://pi.pardot.com/api/prospect/version/4/do/update/id/"+id+"?list_1933=1&api_key=167562362d23c6d60e064d4b86aedc5b&user_key=074fba8df5a2b4d2ab9f244a8a314861";
	 BufferedReader rd1 = new BufferedReader(new InputStreamReader(((HttpURLConnection) (new URL(datas)).openConnection()).getInputStream(), Charset.forName("UTF-8")));
	 String lines;
     while ((lines = rd1.readLine()) != null) {
    	 //System.out.println(lines);
     }
     rd1.close();
     System.out.println(lines);*/

	System.out.println("records inserted row no:"+s);
	System.out.println("-------------------");
	}
	} catch(Exception e) {
		System.out.println("exception is:"+e.getMessage());
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

	