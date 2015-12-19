package Transform;

//import java.io.File;
import java.io.File;
//import java.io.FileInputStream;
import java.io.FileOutputStream;
//import java.io.FileWriter;
import java.io.IOException;
//import java.io.InputStream;
//import java.time.format.DateTimeFormatter;
//import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
//import java.util.GregorianCalendar;
import java.util.HashMap;
//import java.util.List;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.EncryptedDocumentException;
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
//import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;
//import org.apache.poi.xssf.eventusermodel.XSSFReader;
//import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.ss.util.CellReference;
//import org.apache.poi.ss.util.WorkbookUtil;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.*;
//import org.apache.xmlbeans.*;
//import org.xml.sax.ContentHandler;
//import org.xml.sax.SAXException;
//import org.xml.sax.XMLReader;
//import org.xml.sax.helpers.XMLReaderFactory;


public class MonthtoWeek {
	private static ArrayList<String> names;
	private static ArrayList<Integer> Jan;
	private static ArrayList<Integer> Feb;
	private static ArrayList<Integer> Mar;
	private static ArrayList<Integer> Apr;
	private static ArrayList<Integer> May;
	private static ArrayList<Integer> Jun;
	private static ArrayList<Integer> Jul;
	private static ArrayList<Integer> Aug;
	private static ArrayList<Integer> Sep;
	private static ArrayList<Integer> Oct;
	private static ArrayList<Integer> Nov;
	private static ArrayList<Integer> Dec;
	private static ArrayList<Integer> Total;
	private static ArrayList<ArrayList<Integer>> year;
	
	private static int columnbase = 1;
	private static int rowbase = 3;
	private static int machinenamecolumn = 0;
	private static int titlerow = 2;
	private static int titlecolumn = 1;
	
	private static int yearcalendar=2016;
	private static int columnindex = 1;
	private static int weekindex = 1;
	
	public static void main(String[] args) throws IOException, EncryptedDocumentException, OpenXML4JException {
		// TODO Auto-generated method stub
		
		String filename = args[0]; 
		String sheetname = args[1]; 
		try{
			yearcalendar = Integer.parseInt(args[2]);
		}catch(Exception e){
			e.printStackTrace();
		}
		String outputfilename = "./WeeklyView.xlsx";
		
		long enterTime = System.currentTimeMillis();
		System.out.println("Input file name have been recorded");
		MonthtoWeek monthtoweek = new MonthtoWeek();
		
		Jan = new ArrayList<Integer>();
		Feb = new ArrayList<Integer>();
		Mar = new ArrayList<Integer>();
		Apr = new ArrayList<Integer>();
		May = new ArrayList<Integer>();
		Jun = new ArrayList<Integer>();
		Jul = new ArrayList<Integer>();
		Aug = new ArrayList<Integer>();
		Sep = new ArrayList<Integer>();
		Oct = new ArrayList<Integer>();
		Nov = new ArrayList<Integer>();
		Dec = new ArrayList<Integer>();
		Total = new ArrayList<Integer>();
		year = new ArrayList<ArrayList<Integer>>();
		year.add(Jan);
		year.add(Feb);
		year.add(Mar);
		year.add(Apr);
		year.add(May);
		year.add(Jun);
		year.add(Jul);
		year.add(Aug);
		year.add(Sep);
		year.add(Oct);
		year.add(Nov);
		year.add(Dec);
		year.add(Total);

		//input data
		File file = new File(filename);
		Workbook workbook = WorkbookFactory.create(file);
		Sheet monthsheet = workbook.getSheet(sheetname);
		System.out.println("workbook opened");
		monthtoweek.inputdata(monthsheet);
		workbook.close();
		System.out.println("input data completed.");
	
		//output the processed data and release recourses
		Workbook wb = new XSSFWorkbook();
		monthtoweek.outputdata(wb);
		File outputfile = new File(outputfilename);
		FileOutputStream fileout = new FileOutputStream(outputfile);
		wb.write(fileout);
		wb.close();
		fileout.close();
		System.out.println("Work Completed!");
		
		//show the elapsed time
		long leaveTime = System.currentTimeMillis();
		double differencetime = (leaveTime - enterTime)/1000.0;
		System.out.println("Time elapsed: "+differencetime + "seconds");
	}
	
	
	//input data function
	public void inputdata(Sheet monthsheet){
		if(monthsheet == null){
			System.out.println("Input sheet is not valid. Check the sheet's name or make sure it's exsiting.");
			return;
		}
		String pattern = "(Demand\\sPlan\\s)";
		//String patterntest = "Demand\\sPlan\\s(#):\\s20[1-9][0-9]-\\s[01]?[0-9]";
		Pattern titlepattern = Pattern.compile(pattern);
		CellReference cellReference = new CellReference(titlerow,titlecolumn);
		Row cellrow = monthsheet.getRow(cellReference.getRow());
		Cell titlecell = cellrow.getCell(cellReference.getCol());
		if(!(titlecell.getCellType() == Cell.CELL_TYPE_STRING)){
			System.out.println("The table is not found.");
			return;
		}
		
		String titlestring = titlecell.getStringCellValue().trim();
		Matcher m = titlepattern.matcher(titlestring);
		//
		if(m.find()){
			names = new ArrayList<String>();
			int row = rowbase;
			while(true){
				CellReference machineReference = new CellReference(row,machinenamecolumn);
				Row machinerow = monthsheet.getRow(machineReference.getRow());
				Cell machine = machinerow.getCell(machineReference.getCol());
				if(!(machine.getCellType() == Cell.CELL_TYPE_STRING)){
					System.out.println("Machine column is not valid.");
					return;
				}
				String name = machine.getStringCellValue().trim();
				if(name.equals(new String("Total"))){
					break;
				}
				names.add(name);
				row++;
			}
			//add machine name done
			
			//add int number of monthes and total
			for(int j = 0;j<year.size();j++){
				for(int rowofmonth =0;rowofmonth < names.size();rowofmonth++){
					
					CellReference monthReference = new CellReference(rowbase+rowofmonth,columnbase+j);
					Row monthrow = monthsheet.getRow(monthReference.getRow());
					Cell monthcell = monthrow.getCell(monthReference.getCol(),Row.RETURN_NULL_AND_BLANK);
					if ((monthcell == null) || (monthcell.equals("")) || (monthcell.getCellType() == monthcell.CELL_TYPE_BLANK)){
						year.get(j).add(0);
						continue;
					}
					if(!(monthcell.getCellType() == Cell.CELL_TYPE_NUMERIC)){
						System.out.println("Number area in the table is not valid.");
						return;
					}
					int intjannumber = (int) monthcell.getNumericCellValue();
					year.get(j).add(intjannumber);
				}
			}
		}else{
			System.out.println("Sorry, can't find Demand Plan table in current sheet.");
			return;
		}
		return;
	}

	//output function
	public void outputdata(Workbook wb) throws IOException{
		String safename = WorkbookUtil.createSafeSheetName("Consolidation");
		Sheet sheet = wb.createSheet(safename);
		//create table title row
		Row row = sheet.createRow(0);
		//create month title row
		Row monthrow = sheet.createRow(1);
		//create week title row
		Row weekrow = sheet.createRow(2);
		//add machine name column
		Row[] rownames = new Row[names.size()];
		//Row rownames;
		setname(rownames, names, sheet, wb);

		//add month row
		HashMap<Integer, Integer> monthtorange = new HashMap<Integer,Integer>();
		//add week and its cooresponding data and setmonth() is to set month labels
		setweek(year, setmonth(monthtorange, wb, sheet, monthrow),wb, sheet, weekrow);
		//set total column
		settotal(wb, sheet, Total);

		//add title
		settitle(wb,row,sheet);

	}
	
	public void settitle(Workbook wb, Row row, Sheet sheet){
		XSSFCellStyle yeartyle = (XSSFCellStyle) wb.createCellStyle();
	    Font yearfont = wb.createFont();
	    yearfont.setFontHeightInPoints((short)11);
	    yearfont.setFontName("Arial Black");
		
	    yeartyle.setFillBackgroundColor(new XSSFColor(new java.awt.Color(170, 170, 170)));
	    yeartyle.setAlignment(CellStyle.ALIGN_CENTER);
	    yeartyle.setFont(yearfont);
		sheet.addMergedRegion(new CellRangeAddress(0,0,1,columnindex));
		Cell titlecell = row.createCell(1);
		titlecell.setCellValue(yearcalendar + " (CDD)");
		titlecell.setCellStyle(yeartyle);
		return;		
	}
	//set the total column
	public void settotal(Workbook wb, Sheet sheet, ArrayList<Integer> Total){
		for(int i=0; i<Total.size();i++){
			if(i==0){
				
				CellStyle totalstyle = wb.createCellStyle();
				Font totalfont = wb.createFont();
				totalfont.setFontHeightInPoints((short)8);
				totalfont.setFontName("Arial Black");
				totalstyle.setFont(totalfont);
			    totalstyle.setAlignment(CellStyle.ALIGN_CENTER);
			    totalstyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
				
				Row totalr = CellUtil.getRow(1, sheet);
				Cell totalc = CellUtil.getCell(totalr, columnindex);
				totalc.setCellType(Cell.CELL_TYPE_STRING);
			    totalc.setCellValue("Total");
			    totalc.setCellStyle(totalstyle);
			    sheet.addMergedRegion(new CellRangeAddress(1,2,columnindex,columnindex));
			}
			
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setFontName("Calibri");
			font.setFontHeightInPoints((short)8);
			style.setAlignment(CellStyle.ALIGN_CENTER);
			style.setBorderBottom(CellStyle.BORDER_DOTTED);
			style.setBorderRight(CellStyle.BORDER_MEDIUM_DASHED);
			style.setBorderTop(CellStyle.BORDER_DOTTED);
			style.setFont(font);
			
			Row r = CellUtil.getRow(3+i, sheet);
			Cell c = CellUtil.getCell(r, columnindex);
			c.setCellType(Cell.CELL_TYPE_NUMERIC);
			c.setCellValue(Total.get(i)); 
			c.setCellStyle(style);
		}
	}
	
	//set names to output file
	public void setname(Row[] rownames, ArrayList<String> names, Sheet sheet, Workbook wb){
		CellStyle namestyle = wb.createCellStyle();
	    Font namefont = wb.createFont();
		namefont.setFontHeightInPoints((short)10);
		
	    namestyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	    namestyle.setAlignment(CellStyle.ALIGN_RIGHT);
	    namestyle.setFillPattern(CellStyle.NO_FILL);
	    namestyle.setBorderBottom(CellStyle.BORDER_DOTTED);
	    namestyle.setBorderTop(CellStyle.BORDER_DOTTED);
	    namestyle.setBorderRight(CellStyle.BORDER_DASHED);
	    namestyle.setFont(namefont);
	    
		for(int i = 0; i < names.size(); i++){
			rownames[i] = sheet.createRow((short)i+3);
			Cell machine = rownames[i].createCell((short)0);
			machine.setCellType(Cell.CELL_TYPE_STRING);
			machine.setCellValue(names.get(i).toString());
			machine.setCellStyle(namestyle);
			sheet.autoSizeColumn(0, false);
		}
	}
	
	//calculate how many fridays in one month
	public int getweeks(int j){
	    Calendar calendar = Calendar.getInstance();
	    // Note that month is 0-based in calendar, bizarrely.
	    calendar.set(yearcalendar, j, 1);
	    int daysInMonth = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);

	    int count = 0;
	    for (int day = 1; day <= daysInMonth; day++) {
	        calendar.set(yearcalendar, j, day);
	        int dayOfWeek = calendar.get(Calendar.DAY_OF_WEEK);
	        if (dayOfWeek == Calendar.FRIDAY) {
	            count++;
	            
	        }
	    }
	    return count;
	}
	
	//set week labels
	public void addweek(Row weekrow, int range, int start, Workbook wb, Sheet sheet){
		CellStyle weekstyle = wb.createCellStyle();
	    Font weekfont = wb.createFont();
	    weekfont.setFontHeightInPoints((short)8);
	    weekfont.setFontName("Arial Black");
		
	    weekstyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	    weekstyle.setAlignment(CellStyle.ALIGN_CENTER);
	    //weekstyle.setFillPattern(CellStyle.);
	    //weekstyle.setBorderBottom(CellStyle.BORDER_THIN);
	    weekstyle.setBorderTop(CellStyle.BORDER_THIN);
	    weekstyle.setBorderRight(CellStyle.BORDER_THIN);
	    weekstyle.setBorderLeft(CellStyle.BORDER_THIN);
	    weekstyle.setFont(weekfont);
	    
		for(int i = 0; i < range;i++){
			//int week = i+1;
			Cell w = weekrow.createCell(start+i);
			w.setCellType(Cell.CELL_TYPE_STRING);
			w.setCellValue("Week "+ columnindex);
			
			w.setCellStyle(weekstyle);
			sheet.autoSizeColumn(columnindex, false);
			columnindex++;
		}
		return;
	}
	
	//put in data into output file
	public void adddata(Workbook wb, Sheet sheet, int range, ArrayList<Integer> month){

		for(int i=0;i<month.size();i++){
			int numberperweek = month.get(i)/range;
			int remainder = month.get(i)%range;
			Random generator = new Random(); 
			//System.out.print(i);
			int randomint = generator.nextInt(range);
			//add data to every week in this month every machine
			for(int j = 0; j<range;j++){
				Row r = CellUtil.getRow(3+i, sheet);
				//System.out.println("r's row: "+ ref.getRow());
				if (r != null) {
					Cell c = CellUtil.getCell(r, weekindex+j);
					c.setCellType(Cell.CELL_TYPE_NUMERIC);
					CellStyle style = wb.createCellStyle();
					Font font = wb.createFont();
					font.setFontName("Calibri");
					font.setFontHeightInPoints((short)8);
					style.setAlignment(CellStyle.ALIGN_CENTER);
					style.setBorderBottom(CellStyle.BORDER_DOTTED);
					if(i!=0){
						style.setBorderTop(CellStyle.BORDER_DOTTED);
					}else{
						style.setBorderTop(CellStyle.BORDER_THIN);
					}
					
					style.setFont(font);
					if(j!=0){
						style.setBorderLeft(CellStyle.BORDER_DOTTED);
					}//else{
					//	style.setBorderLeft(CellStyle.BORDER_MEDIUM_DASHED);
					//}
					if(j==range-1){
						style.setBorderRight(CellStyle.BORDER_MEDIUM_DASHED);
					}else{
						style.setBorderRight(CellStyle.BORDER_DOTTED);
					}
					c.setCellStyle(style);
				    if(randomint == j){
				    	c.setCellValue((double)(numberperweek+remainder));
				    	//System.out.print((numberperweek+remainder));
				    }else{
				    	c.setCellValue((double)(numberperweek));
				    	//System.out.print((numberperweek));
				    }
				 }
			}
			if(i==month.size()-1){
				weekindex = weekindex+range;
				//System.out.println("columnindex: "+ columnindex);
			}
		}
		return;
	}
	
	public void setweek(ArrayList<ArrayList<Integer>> year, int ranges[], Workbook wb, Sheet sheet, Row weekrow){
		int startinsep = ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+ranges[7]+1;
		int endinsep = ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+ranges[7]+ranges[8];
		for(int j =0; j<year.size()-1;j++){
			if(j==0){
				addweek(weekrow,ranges[j],1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==1){
				addweek(weekrow,ranges[j],ranges[0]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==2){
				addweek(weekrow,ranges[j],ranges[0]+ranges[1]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==3){
				addweek(weekrow,ranges[j],ranges[0]+ranges[1]+ranges[2]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==4){
				addweek(weekrow,ranges[j],ranges[0]+ranges[1]+ranges[2]+ranges[3]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==5){
				addweek(weekrow,ranges[j],ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==6){
				addweek(weekrow,ranges[j],ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==7){
				addweek(weekrow,ranges[j],ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==8){
				addweek(weekrow,ranges[j],startinsep,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==9){
				addweek(weekrow,ranges[j],endinsep+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==10){
				addweek(weekrow,ranges[j],endinsep+ranges[9]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
			if(j==11){
				addweek(weekrow,ranges[j],endinsep+ranges[9]+ranges[10]+1,wb,sheet);
				adddata(wb,sheet,ranges[j], year.get(j));
			}
		}
	}
	
	public int[] setmonth(HashMap<Integer, Integer> monthtorange, Workbook wb, Sheet sheet, Row monthrow){
		int ranges[] = new int[12];
		CellStyle monthstyle = wb.createCellStyle();
	    Font monthfont = wb.createFont();
	    monthfont.setFontHeightInPoints((short)8);
		
		monthstyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		monthstyle.setAlignment(CellStyle.ALIGN_CENTER);
		//monthstyle.setFillPattern(CellStyle.NO_FILL);
		monthstyle.setBorderBottom(CellStyle.BORDER_THIN);
		//monthstyle.setBorderTop(CellStyle.BORDER_THIN);
		monthstyle.setBorderRight(CellStyle.BORDER_THIN);
		monthstyle.setBorderLeft(CellStyle.BORDER_THIN);
		monthstyle.setFont(monthfont);    
	    
		Cell monthcell = monthrow.createCell(1);
		monthcell.setCellType(Cell.CELL_TYPE_STRING);
		monthcell.setCellValue("Jan");
		ranges[0] = getweeks(0);
		monthtorange.put(0,ranges[0]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,1,ranges[0]));
		monthcell.setCellStyle(monthstyle);
		
		Cell monthcell1 = monthrow.createCell(ranges[0]+1);
		monthcell1.setCellType(Cell.CELL_TYPE_STRING);
		monthcell1.setCellValue("Feb");
		ranges[1] = getweeks(1);
		monthtorange.put(1,ranges[1]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,ranges[0]+1,ranges[0]+ranges[1]));
		monthcell1.setCellStyle(monthstyle);
		
		Cell monthcell2 = monthrow.createCell(ranges[0]+ranges[1]+1);
		monthcell2.setCellType(Cell.CELL_TYPE_STRING);
		monthcell2.setCellValue("Mar");
		ranges[2] = getweeks(2);
		monthtorange.put(2,ranges[2]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,ranges[0]+ranges[1]+1,ranges[0]+ranges[1]+ranges[2]));
		monthcell2.setCellStyle(monthstyle);

		Cell monthcell3 = monthrow.createCell(ranges[0]+ranges[1]+ranges[2]+1);
		monthcell3.setCellType(Cell.CELL_TYPE_STRING);
		monthcell3.setCellValue("Apr");
		ranges[3] = getweeks(3);
		monthtorange.put(3,ranges[3]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,ranges[0]+ranges[1]+ranges[2]+1,ranges[0]+ranges[1]+ranges[2]+ranges[3]));
		monthcell3.setCellStyle(monthstyle);
		
		Cell monthcell4 = monthrow.createCell(ranges[0]+ranges[1]+ranges[2]+ranges[3]+1);
		monthcell4.setCellType(Cell.CELL_TYPE_STRING);
		monthcell4.setCellValue("May");
		ranges[4] = getweeks(4);
		monthtorange.put(4,ranges[4]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]));
		monthcell4.setCellStyle(monthstyle);

		Cell monthcell5 = monthrow.createCell(ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+1);
		monthcell5.setCellType(Cell.CELL_TYPE_STRING);
		monthcell5.setCellValue("Jun");
		ranges[5] = getweeks(5);
		monthtorange.put(5,ranges[5]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]));
		monthcell5.setCellStyle(monthstyle);
		
		Cell monthcell6 = monthrow.createCell(ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+1);
		monthcell6.setCellType(Cell.CELL_TYPE_STRING);
		monthcell6.setCellValue("Jul");
		ranges[6] = getweeks(6);
		monthtorange.put(6,ranges[6]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]));
		monthcell6.setCellStyle(monthstyle);
		
		Cell monthcell7 = monthrow.createCell(ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+1);
		monthcell7.setCellType(Cell.CELL_TYPE_STRING);
		monthcell7.setCellValue("Aug");
		ranges[7] = getweeks(7);
		monthtorange.put(7,ranges[7]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+1,ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+ranges[7]));
		monthcell7.setCellStyle(monthstyle);
		
		int startinsep = ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+ranges[7]+1;
		
		Cell monthcell8 = monthrow.createCell(startinsep);
		monthcell8.setCellType(Cell.CELL_TYPE_STRING);
		monthcell8.setCellValue("Sep");
		ranges[8] = getweeks(8);
		monthtorange.put(8,ranges[8]);
		int endinsep = ranges[0]+ranges[1]+ranges[2]+ranges[3]+ranges[4]+ranges[5]+ranges[6]+ranges[7]+ranges[8];
		sheet.addMergedRegion(new CellRangeAddress(1,1,startinsep,endinsep));
		monthcell8.setCellStyle(monthstyle);
		
		Cell monthcell9 = monthrow.createCell(endinsep+1);
		monthcell9.setCellType(Cell.CELL_TYPE_STRING);
		monthcell9.setCellValue("Oct");
		ranges[9] = getweeks(9);
		monthtorange.put(9,ranges[9]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,endinsep+1,endinsep+ranges[9]));
		monthcell9.setCellStyle(monthstyle);
		
		Cell monthcell10 = monthrow.createCell(endinsep+ranges[9]+1);
		monthcell10.setCellType(Cell.CELL_TYPE_STRING);
		monthcell10.setCellValue("Nov");
		ranges[10] = getweeks(10);
		monthtorange.put(10,ranges[10]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,endinsep+ranges[9]+1,endinsep+ranges[9]+ranges[10]));
		monthcell10.setCellStyle(monthstyle);
		
		Cell monthcell11 = monthrow.createCell(endinsep+ranges[9]+ranges[10]+1);
		monthcell11.setCellType(Cell.CELL_TYPE_STRING);
		monthcell11.setCellValue("Dec");
		ranges[11] = getweeks(11);
		monthtorange.put(11,ranges[11]);
		sheet.addMergedRegion(new CellRangeAddress(1,1,endinsep+ranges[9]+ranges[10]+1,endinsep+ranges[9]+ranges[10]+ranges[11]));
		monthcell11.setCellStyle(monthstyle);
		
		return ranges; 
	}
}
