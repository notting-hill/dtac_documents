import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.commons.io.FilenameUtils;

public class MatchingOriginal {
	final static String prefixTable = "DIM_CCB_";

	public static void main(String[] args) throws Exception {
		if(args.length > 0 && "create".equals(args[0])) {
			prepareFiles();
			System.exit(0);
		} else {
			File dir = new File(".");
			File[] buFiles = dir.listFiles(new FilenameFilter() {
				@Override
					public boolean accept(File dir, String name) {
						return name.startsWith("bu_") && name.endsWith(".txt");
					}
			});

			File outFile = new File("output.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheetSummary = workbook.createSheet("Summary");

			for(int i=0; i<buFiles.length; i++) {
				File buFile = buFiles[i];
				String feedName = buFile.getName().substring(3, buFile.getName().length()-4);
				File stmFile = new File("stm_" + feedName + ".txt");
				System.out.println("process " + buFile.getName() + "," + stmFile.getName() + " generating: " + outFile.getName());
				process2(buFile, stmFile, workbook, sheetSummary);
			}

		   try (FileOutputStream outputStream = new FileOutputStream(outFile)) {
				workbook.write(outputStream);
			}
		}
	}

	public synchronized static void process(File buFile, File stmFile, File outFile) throws FileNotFoundException, IOException {
		HashMap h = readSourceFile(buFile);
		// Manage in Target
		BufferedReader br = null;
		String line;
		PrintWriter out = new PrintWriter(new FileOutputStream(outFile, false));
		List sheet = new ArrayList();
		try {
			br = new BufferedReader(new FileReader(stmFile));
			while((line = br.readLine()) != null){
				line = line.replaceAll(",","").replaceAll("\"","").trim();
				String[] field = line.split("\\s+");
				String fieldName = field[0].trim();
				List fieldValues = (List)h.get(fieldName);
				String dataType = ""; 
				String nullable = "";
				dataType = (fieldValues != null && fieldValues.get(0) != null)? (String)(fieldValues.get(0)) : "";
				nullable = (fieldValues != null && fieldValues.get(1) != null)? (String)(fieldValues.get(1)) : "Y";

				if(!"COMP_CD".equals(fieldName) && !"LOAD_DATE".equals(fieldName) && !"LOAD_USER".equals(fieldName) && !"FILE_ID".equals(fieldName)
					&& !"FILE_DATE".equals(fieldName) && !"PROCESS_NAME".equals(fieldName) && !"SBSCRPN_ID".equals(fieldName)
				)
				/*System.*/out.println(padRight(fieldName, 30) + "\t\t" + padRight(mapping(/*fieldName,*/ dataType),30) + "\t\t" + padRight(dataType, 30) + "\t\t" + padRight(nullable, 20));
			}
			//out.println(padRight("COMP_CD",30)+"\t\tVARCHAR2(5)\t\t");
			/*out.println(padRight("LOAD_DATE",30)+"\t\t"+padRight("N",20)+"\t\tDATETIME");
			out.println(padRight("LOAD_USER",30)+"\t\t"+padRight("N",20)+"\t\tVARCHAR2(20)");
			out.println(padRight("FILE_ID",30)+"\t\t"+padRight("N",20)+"\t\tINTEGER");
			out.println(padRight("FILE_DATE",30)+"\t\t"+padRight("N",20)+"\t\tDATETIME");
			out.println(padRight("PROCESS_NAME",30)+"\t\t"+padRight("N",20)+"\t\tVARCHAR2(150)");*/
		} finally {
			if(out != null)
				out.close();
		}
	}


	public synchronized static void process2(File buFile, File stmFile, XSSFWorkbook workbook, XSSFSheet sheetSummary) throws FileNotFoundException, IOException {
		HashMap h = readSourceFile(buFile);
		// Manage in Target
		String buFileNameWithOutExt = FilenameUtils.removeExtension(buFile.getName());
		buFileNameWithOutExt = buFileNameWithOutExt.substring(3,buFileNameWithOutExt.length());

		String line;
		List list = new ArrayList();
		List list2 = new ArrayList();

		list2.add(Arrays.asList(""));
		list2.add(Arrays.asList(buFileNameWithOutExt.toUpperCase()));
		list2.add(Arrays.asList(""));
		list2.add(Arrays.asList("S.NO","Field Name","Data Type","Sample Values"));
		
		BufferedReader br = new BufferedReader(new FileReader(stmFile));
		int rowId = 1;
		while((line = br.readLine()) != null){
			line = line.replace(",","").replace("\"","").trim();
			line = line.replace("*","");
			String[] field = line.split("\\s+");
			String fieldName = field[0].trim();
			List fieldValues = (List)h.get(fieldName);
			String dataType = ""; 
			String nullable = "";
			dataType = (fieldValues != null && fieldValues.get(0) != null)? (String)(fieldValues.get(0)) : "";
			nullable = (fieldValues != null && fieldValues.get(1) != null)? (String)(fieldValues.get(1)) : "Y";

			/*if(!"COMP_CD".equals(fieldName) && !"LOAD_DATE".equals(fieldName) && !"LOAD_USER".equals(fieldName) && !"FILE_ID".equals(fieldName)
				&& !"FILE_DATE".equals(fieldName) && !"PROCESS_NAME".equals(fieldName) && !"SBSCRPN_ID".equals(fieldName)
			)*/
			list.add(Arrays.asList(fieldName,"","","",prefixTable + buFileNameWithOutExt.toUpperCase(),fieldName,mapping(dataType),dataType,"Direct",nullable));
			list2.add(Arrays.asList(new Integer(rowId++), fieldName, mapping(dataType)));
		}

		XSSFSheet sheet1 = workbook.createSheet(buFileNameWithOutExt);
		writeToSheet(sheet1, list);
		writeAppendToSheet(sheetSummary, list2);
	}


	public static HashMap readSourceFile(File f) throws FileNotFoundException, IOException {
		BufferedReader br = null; 
		try {
			br = new BufferedReader(new FileReader(f)); 
			String line;
			HashMap h = new HashMap();
			while((line = br.readLine()) != null){
				String[] field = line.split("\\s+");
				String fieldName = field[0].trim();
				String dataType = field[1].trim();
				String nullable = null;	
				try {
					nullable = field[2].trim();
				} catch (Exception ee) {
					nullable = "";
				}
				h.put(fieldName, Arrays.asList(dataType,nullable));
			}
			return h;
		} finally {
		}
	}

	private static void prepareFiles() {
		File file = new File("tables.txt");
		BufferedReader br = null; 
		try {
			br = new BufferedReader(new FileReader(file)); 
			String line;
			HashMap h = new HashMap();
			while((line = br.readLine()) != null){
				new File("bu_" + line + ".txt").createNewFile();
				new File("stm_" + line + ".txt").createNewFile();
			}
		} catch(Exception e) {
		}
	}

	private static String mapping(/*String fieldName,*/ String type) {
		/*if("SBSCRPN_ID".equals(fieldName))
			return "VARCHAR2(30)";*/
		if(type == null)
			return "";
		if (type.contains("VARCHAR") || type.contains("CHAR")) {
			int firstBranket = type.indexOf("(");
			int lastBranket = type.indexOf(")");
			return String.format("VARCHAR2(%s)", type.substring(firstBranket+1, lastBranket));
		} else if((type.contains("DECIMAL") /*&& type.contains("2")*/)) {
			int firstBranket = type.indexOf("(");
			int firstComma = type.indexOf(",");
			String precision = type.substring(firstBranket+1, firstComma);
			String digit = type.substring(firstComma+1, type.length()-1);
			/*if(Integer.parseInt(precision) > 22)
				return "NUMBER(" + precision + ",2)";
			else if(Integer.parseInt(precision) >= 16)
				return "NUMBER(22,2)";
			else
				return "NUMBER(15,2)";*/
			return "NUMBER(" + precision + "," +digit+ ")";
		} else if("BYTEINT".equals(type) || "BIGINT".equals(type) || "SMALLINT".equals(type) /*|| (type.contains("DECIMAL") && type.contains("0"))*/) {
			//return "INTEGER";
			return type;
		} else if ("TIMESTAMP(0)".equals(type) || "TIMESTAMP(6)".equals(type)) {
			return "DATETIME";
		} /*else if (type.contains("DECIMAL")) {
			if(fieldName.contains("AMT") || fieldName.contains("AMOUNT"))
				return "NUMBER(22,2)";
			return "INTEGER";
		}*/
		return type;
	}

	public static String padRight(String s, int n) {
		 return String.format("%-" + n + "s", s);  
	}

	private static void writeToSheet(XSSFSheet sheet, List datas) {
        int rowCount = 0;
        for (Object rowData : datas) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;
            
            for (Object field : ((List)rowData)) {
                Cell cell = row.createCell(columnCount++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
				} else if (field instanceof Number) {
                    cell.setCellValue(NumberFormat.getInstance(new Locale("en", "US")).format(field));
				} else if (field instanceof Date) {
					cell.setCellValue(new SimpleDateFormat("YYYY-MM-dd hh:mm:ss").format(field));	
				} else {
					cell.setCellValue(field.toString());
				}
            }
        }
	}

	private static void writeAppendToSheet(XSSFSheet sheet, List datas) {
        int rowCount = 0;
		int lastRow = sheet.getLastRowNum();
		rowCount = lastRow + 1;

        for (Object rowData : datas) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;
            
            for (Object field : ((List)rowData)) {
                Cell cell = row.createCell(columnCount++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
				} else if (field instanceof Number) {
                    cell.setCellValue(NumberFormat.getInstance(new Locale("en", "US")).format(field));
				} else if (field instanceof Date) {
					cell.setCellValue(new SimpleDateFormat("YYYY-MM-dd hh:mm:ss").format(field));	
				} else {
					cell.setCellValue(field.toString());
				}
            }
        }
	}

}