import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.commons.io.FilenameUtils;

public class GenerateTable {
	private final static String PREFIX_SOURCE = "CCB";
	private final static String TABLE_FILEFEEDLIST = "FILEFEEDLIST";
	private final static String TABLE_FILETRANSFERMECHANISM1 = "FILETRANSFERMECHANISM1";
	private final static String TABLE_FILETRANSFERMECHANISM2 = "FILETRANSFERMECHANISM2";
	private final static String TABLE_SOURCEFEEDFILES = "SOURCEFEEDFILES";
	private final static String TABLE_FIRSTLOAD = "FIRSTLOAD";
	private final static String TABLE_FEED2TABLE = "FEED2TABLE";
	private final static String TABLE_CONTROLFEED = "CONTROLFEED";
	private final static String TABLE_AVAILABILITY = "AVAILABILITY";
	
	public static void main(String[] args) throws Exception {
		//HashMap<String,List> h = readSourceFile(new File("bu.txt"));
		//printSortedHashmap(h);
		HashMap h = readBuExcel(new File("bu.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet1 = workbook.createSheet(TABLE_FILEFEEDLIST);
        XSSFSheet sheet2 = workbook.createSheet(TABLE_FILETRANSFERMECHANISM1);
        XSSFSheet sheet3 = workbook.createSheet(TABLE_FILETRANSFERMECHANISM2);
        XSSFSheet sheet4 = workbook.createSheet(TABLE_SOURCEFEEDFILES);
        XSSFSheet sheet5 = workbook.createSheet(TABLE_FIRSTLOAD);
        XSSFSheet sheet6 = workbook.createSheet(TABLE_FEED2TABLE);
        XSSFSheet sheet7 = workbook.createSheet(TABLE_CONTROLFEED);
        XSSFSheet sheet8 = workbook.createSheet(TABLE_AVAILABILITY);

		writeToSheet(sheet1, (List)h.get(TABLE_FILEFEEDLIST));          
		writeToSheet(sheet2, (List)h.get(TABLE_FILETRANSFERMECHANISM1));
		writeToSheet(sheet3, (List)h.get(TABLE_FILETRANSFERMECHANISM2));
		writeToSheet(sheet4, (List)h.get(TABLE_SOURCEFEEDFILES));       
		writeToSheet(sheet5, (List)h.get(TABLE_FIRSTLOAD));             
		writeToSheet(sheet6, (List)h.get(TABLE_FEED2TABLE));            
		writeToSheet(sheet7, (List)h.get(TABLE_CONTROLFEED));           
		writeToSheet(sheet8, (List)h.get(TABLE_AVAILABILITY));          

        try (FileOutputStream outputStream = new FileOutputStream("out_"+PREFIX_SOURCE+".xlsx")) {
            workbook.write(outputStream);
        }
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


	/*public static void readExcel(File file) throws FileNotFoundException, IOException {
		FileInputStream excelFile = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet datatypeSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = datatypeSheet.iterator();		
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				if (currentCell.getCellType() == CellType.STRING) {
					System.out.print(currentCell.getStringCellValue() + "--");
				} else if (currentCell.getCellType() == CellType.NUMERIC) {
					System.out.print(currentCell.getNumericCellValue() + "--");
				}
			}
			System.out.println();
		}
	}*/

	public static HashMap readBuExcel(File xlsxFile) throws FileNotFoundException, IOException {
		HashMap h = new HashMap();
		List lFILEFEEDLIST = new ArrayList();
		lFILEFEEDLIST.add(Arrays.asList(new String[]{"FEEDNAME","FILENAME","FILEEXT","CONTACT","CONTACTEMAIL","FILEDESC"}));
		List lFILETRANSFERMECHANISM1 = new ArrayList();
		lFILETRANSFERMECHANISM1.add(Arrays.asList(new String[]{"FEEDNAME","FILENAME","FILEEXT","CONTROLFILENAME","CONTROLFILEEXT"}));
		List lFILETRANSFERMECHANISM2 = new ArrayList();
		lFILETRANSFERMECHANISM2.add(Arrays.asList(new String[]{"FEEDNAME","FILE","PRODSOURCEFILEPATH"/*,"SITSOURCEFILEPATH","DEVSOURCEFILEPATH"*/}));
		List lSOURCEFEEDFILES = new ArrayList();
		lSOURCEFEEDFILES.add(Arrays.asList(new String[]{"FEEDNAME","FILENAME","FILEEXT","CHARSET","CONV","FILETYPE","FREQ","PULL","MINNOFILES","DELIMITER","INTEGRATIONTYPE","HEADER"}));
		List lFIRSTLOAD = new ArrayList();
		lFIRSTLOAD.add(Arrays.asList(new String[]{"FEEDNAME","FILENAME","FILEEXT","HISTORICALLOAD","6 MONTHS"}));
		List lFEED2TABLE = new ArrayList();
		lFEED2TABLE.add(Arrays.asList(new String[]{"FEEDNAME","TABLENAME","FILESIZE","VOLUMEPERDAY","FILETYPE","FILEDATEREF"}));
		List lCONTROLFEED = new ArrayList();
		lCONTROLFEED.add(Arrays.asList(new String[]{"FEEDNAME","FILENAME","CONTROLFILENAME","CONTROLFORMAT"}));
		List lAVAILABILITY = new ArrayList();
		lAVAILABILITY.add(Arrays.asList(new String[]{"FEEDNAME","FILENAME","AVAILABLETIME","AEP_LANDING","CONTROL_PATH"}));

		FileInputStream excelFile = new FileInputStream(xlsxFile);
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet datatypeSheet = workbook.getSheet(PREFIX_SOURCE);
		Iterator<Row> iterator = datatypeSheet.iterator();
		boolean skipHeader = true;
		while (iterator.hasNext()) {
			if(skipHeader) {
				iterator.next();
				skipHeader = false;
			}
			Row row = iterator.next();
			String feedName = row.getCell(CellReference.convertColStringToIndex("N")).toString();
			String source = row.getCell(CellReference.convertColStringToIndex("C")).toString();
			String appName = row.getCell(CellReference.convertColStringToIndex("D")).toString();
			source = ("".equals(source))?"DEF":source;
			feedName = feedName.replaceAll("<YYYYMMDD>_", "").replaceAll("_<YYYYMMDD>", "").toLowerCase();
			feedName = feedName.replaceAll(".txt","").replaceAll(".csv","").replaceAll(".dat","");
			feedName = feedName.toLowerCase();
			if(appName != null)
				feedName = appName.toLowerCase();
			String fileType = row.getCell(CellReference.convertColStringToIndex("L")).toString();
			String targetTable = (("MASTER".equalsIgnoreCase(fileType))?"DIM":"DWO") + "_" + PREFIX_SOURCE + "_" + feedName;
			targetTable = targetTable.toUpperCase();
			String fileDesc = row.getCell(CellReference.convertColStringToIndex("X")).toString();

			String processName = row.getCell(CellReference.convertColStringToIndex("E")).toString();
			//if(processName.indexOf("_DTN") != -1)
				//feedName = feedName + "_dtn";
			String fileName = row.getCell(CellReference.convertColStringToIndex("AE")).toString();
			String fileNamePattern = row.getCell(CellReference.convertColStringToIndex("J")).toString();
			String fileNameWithoutExt = FilenameUtils.removeExtension(fileNamePattern);
			System.out.println(fileNameWithoutExt);
			if(!"".equals(fileName))
				fileName = fileName.substring(fileName.lastIndexOf("/")+1, fileName.length());
			String fileNameWithExample = (!fileName.contains("2019"))?fileName:(fileNameWithoutExt + "(ex. " + fileName + ")");
			if(fileNameWithExample == null || "".equals(fileNameWithExample))
				fileNameWithExample = fileNameWithoutExt;

			String fileExt = row.getCell(CellReference.convertColStringToIndex("K")).toString();
			String path = row.getCell(CellReference.convertColStringToIndex("H")).toString().replaceAll("/PROD/EDW/SRC_DATA","").replaceAll("/PROD/EDW/","").replaceAll("/","");

			String contact = row.getCell(CellReference.convertColStringToIndex("P")).toString();
			String contactEmail = row.getCell(CellReference.convertColStringToIndex("Q")).toString();

			String sourceDataPath = row.getCell(CellReference.convertColStringToIndex("AK")).toString();
			String prodSourceFilePath = row.getCell(CellReference.convertColStringToIndex("AP")).toString();

			String sourcePathFor = row.getCell(CellReference.convertColStringToIndex("AL")).toString();
			//String sitSourceFilePath = sourcePathFor.replaceAll("<root>","SIT/EDW/SRC_DATA").trim();
			//String devSourceFilePath = sourcePathFor.replaceAll("<root>","DEV/EDW/SRC_DATA").trim();

			String freq = row.getCell(CellReference.convertColStringToIndex("R")).toString();
			String freqSource = row.getCell(CellReference.convertColStringToIndex("S")).toString();
			String minNoFiles = row.getCell(CellReference.convertColStringToIndex("T")).toString();
			String delimiter = row.getCell(CellReference.convertColStringToIndex("U")).toString();
			String integrationType = row.getCell(CellReference.convertColStringToIndex("AQ")).toString();
			if("".equals(integrationType))
				integrationType = "Increment";
			String historicalLoad = row.getCell(CellReference.convertColStringToIndex("AT")).toString();
			String fileDateRef1 = row.getCell(CellReference.convertColStringToIndex("AX")).toString();
			String fileDateRef2 = row.getCell(CellReference.convertColStringToIndex("AY")).toString();
			String fileDateRef = "";
			if ("Near real time".equalsIgnoreCase(fileDateRef2))
				fileDateRef = "01";
			else if(fileNamePattern.indexOf("YYYY") == -1)
				fileDateRef = "02";
			else if(fileNamePattern.indexOf("YYYY") != -1 && "Today".equals(fileDateRef1) && "Today -1".equals(fileDateRef2))
				fileDateRef = "03";
			else if(fileNamePattern.indexOf("YYYY") != -1 && "Today -1".equals(fileDateRef2))
				fileDateRef = "04";

			String headerFlag = row.getCell(CellReference.convertColStringToIndex("V")).toString();
			double recordPerDay = 0;
			double sizeOnBytes = 0;
			try { sizeOnBytes = (row.getCell(CellReference.convertColStringToIndex("AM"))).getNumericCellValue(); } catch(Exception e) {}
			try { recordPerDay = (row.getCell(CellReference.convertColStringToIndex("AN"))).getNumericCellValue(); } catch(Exception e) {}

			String availableTime = getCellValueAsString(row.getCell(CellReference.convertColStringToIndex("BA")));

			boolean hasControlFile = (row.getCell(CellReference.convertColStringToIndex("AG")).toString().startsWith("Y"))? true: false;
			String controlFormat = row.getCell(CellReference.convertColStringToIndex("AH")).toString();
			String controlFileName = "";
			String controlFileExt = "";
			String controlFilePath = "";

			controlFilePath = row.getCell(CellReference.convertColStringToIndex("AP")).toString();
			try { controlFileName = controlFilePath.substring(controlFilePath.lastIndexOf("/")+1, controlFilePath.length()); } catch(Exception e) {}
			try { controlFileExt = controlFileName.substring(controlFileName.indexOf("."), controlFileName.length()); } catch(Exception e) {}
		

			String aepDataPath = "";
			String aepControlPath = "";
			String landingPath = row.getCell(CellReference.convertColStringToIndex("AL")).toString();
			try { aepDataPath = "<<Mount_point>>/AEP_landing" + landingPath.replaceAll("/<root>",""); } catch(Exception e) {}
			try { aepControlPath = "<<Mount_point>>/AEP_landing" + landingPath.replaceAll("/<root>","") + "/CTRL/";  } catch(Exception e) {};

			//String test = row.getCell(CellReference.convertColStringToIndex("F")).toString();
			//System.out.println(test);

			lFILEFEEDLIST.add(Arrays.asList(feedName,fileNameWithoutExt,fileExt,contact,contactEmail,fileDesc));
			lFILETRANSFERMECHANISM1.add(Arrays.asList(feedName,fileNameWithoutExt,fileExt,controlFileName,controlFileExt));
			lFILETRANSFERMECHANISM2.add(Arrays.asList(feedName,fileNameWithExample,sourceDataPath + fileNameWithExample/*,sitSourceFilePath,devSourceFilePath*/));
			lSOURCEFEEDFILES.add(Arrays.asList(feedName,fileNameWithoutExt,fileExt,"","",fileType,freq,"Pull",minNoFiles,delimiter,integrationType,headerFlag));
			lFIRSTLOAD.add(Arrays.asList(feedName,fileNameWithoutExt,fileExt,historicalLoad,"6 Months"));
			lFEED2TABLE.add(Arrays.asList(feedName,targetTable,sizeOnBytes,recordPerDay,fileType,fileDateRef));
			lCONTROLFEED.add(Arrays.asList(feedName,fileName,controlFileName,controlFormat));
			lAVAILABILITY.add(Arrays.asList(feedName,fileNameWithExample,availableTime,landingPath,landingPath));

			/*Iterator<Cell> cellIterator = row.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				if (currentCell.getCellType() == CellType.STRING) {
					System.out.print(currentCell.getStringCellValue() + "--");
				} else if (currentCell.getCellType() == CellType.NUMERIC) {
					System.out.print(currentCell.getNumericCellValue() + "--");
				}
			}*/
			//System.out.println();
		}
		h.put(TABLE_FILEFEEDLIST, lFILEFEEDLIST);
		h.put(TABLE_FILETRANSFERMECHANISM1, lFILETRANSFERMECHANISM1);
		h.put(TABLE_FILETRANSFERMECHANISM2, lFILETRANSFERMECHANISM2);
		h.put(TABLE_SOURCEFEEDFILES, lSOURCEFEEDFILES);
		h.put(TABLE_FIRSTLOAD, lFIRSTLOAD);
		h.put(TABLE_FEED2TABLE, lFEED2TABLE);
		h.put(TABLE_CONTROLFEED, lCONTROLFEED);
		h.put(TABLE_AVAILABILITY, lAVAILABILITY);
		return h;
	}

	private static String getCellValueAsString(Cell poiCell){
		if (poiCell.getCellType()==CellType.NUMERIC && DateUtil.isCellDateFormatted(poiCell)) {
			//get date
			Date date = poiCell.getDateCellValue();

			//set up formatters that will be used below
			SimpleDateFormat formatTime = new SimpleDateFormat("HH:mm:ss");
			SimpleDateFormat formatYearOnly = new SimpleDateFormat("yyyy");

			/*get date year.
			*"Time-only" values have date set to 31-Dec-1899 so if year is "1899"
			* you can assume it is a "time-only" value 
			*/
			String dateStamp = formatYearOnly.format(date);

			if (dateStamp.equals("1899")){
				//Return "Time-only" value as String HH:mm:ss
				return formatTime.format(date);
			} else {
				//here you may have a date-only or date-time value

				//get time as String HH:mm:ss 
				String timeStamp =formatTime.format(date);

				if (timeStamp.equals("00:00:00")){
					//if time is 00:00:00 you can assume it is a date only value (but it could be midnight)
					//In this case I'm fine with the default Cell.toString method (returning dd-MMM-yyyy in case of a date value)
					return poiCell.toString();
				} else {
					//return date-time value as "dd-MMM-yyyy HH:mm:ss"
					return poiCell.toString()+" "+timeStamp;
				}
			}
		}
		//use the default Cell.toString method (returning "dd-MMM-yyyy" in case of a date value)
		return poiCell.toString();
	}

	/*private static void writeToFile(Object[][] datas, String sheetName, int startRowIdx, int startColIdx) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        int rowCount = 0 + startRowIdx;
        for (Object[] rowData : datas) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0 + startColIdx;
             
            for (Object field : rowData) {
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
        try (FileOutputStream outputStream = new FileOutputStream("out.xlsx")) {
            workbook.write(outputStream);
        }
	}*/

	/*private static void writeToFile(List datas, String fileName, String sheetName, int startRowIdx, int startColIdx) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        int rowCount = 0 + startRowIdx;
        for (Object rowData : datas) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0 + startColIdx;
            
            for (Object field : ((Object[])rowData)) {
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
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
        }
	}*/

	/*public static HashMap<String,List> readSourceFile(File f) throws FileNotFoundException, IOException {
		BufferedReader br = null; 
		try {
			br = new BufferedReader(new FileReader(f)); 
			String line;
			HashMap h = new HashMap();
			while((line = br.readLine()) != null){
				String[] fields = line.split("\\s+");
				String fieldName = fields[indexForColumn("N")].trim();
				String processName = fields[indexForColumn("E")].trim();
				String fileName = fields[indexForColumn("J")].trim();
				String fileExt = fields[indexForColumn("K")].trim();
				String file = fileName + fileExt;
				String contact = fields[indexForColumn("P")].trim();
				String contactEmail = fields[indexForColumn("Q")].trim();
				//System.out.println(fieldName);
				// FILEFEEDLIST
				h.put(fieldName + "_" + TABLE_FILEFEEDLIST, Arrays.asList(fieldName,fileName,fileExt,contact,contactEmail));
				
				// TABLE_FILETRANSFERMECHANISM
				String controlFileName = "";
				String controlFileExt = "";

				String prodSourceFilePath = getFieldValue(fields, "AP");
				String sitSourceFilePath = getFieldValue(fields,"AE").replaceAll("PROD","SIT").trim();
				String devSourceFilePath = getFieldValue(fields,"AE").replaceAll("PROD","DEV").trim();
				boolean hasControlFile = ("Y".equals(getFieldValue(fields,"AG")))? true: false;
				String controlFilePath = getFieldValue(fields,"AP");
				try { controlFileName = controlFilePath.substring(controlFilePath.lastIndexOf("/"), controlFilePath.length()); } catch(Exception e) {}
				try { controlFileExt = controlFileName.substring(controlFileName.indexOf("."), controlFileName.length()); } catch(Exception e) {}
				String fileType = getFieldValue(fields,"L");
				String freq = getFieldValue(fields,"R");
				String freqSource = getFieldValue(fields,"S");
				String minNoFiles = getFieldValue(fields,"T");
				String delimiter = getFieldValue(fields,"U");
				String integrationType = getFieldValue(fields,"AQ");
				String historicalLoad = getFieldValue(fields,"AT");
				String source = getFieldValue(fields,"D"); 
				String tableName = "DIM_"+source+"_"+processName.replaceAll("P_LD_","");
				String recordPerDay = getFieldValue(fields,"AN"); 
				String controlFormat = getFieldValue(fields,"AH");
				String availableTime = getFieldValue(fields,"AW");

				String aepDataPath = "";
				String aepControlPath = "";
				try { aepDataPath = "<<Mount_point>>/AEP_landing" + getFieldValue(fields,"AL").replaceAll("/<root>/SRC_DATA",""); } catch(Exception e) {}
				try { aepControlPath = "<<Mount_point>>/AEP_landing" + getFieldValue(fields,"AL").replaceAll("/<root>/SRC_DATA","") + "/CTRL/";  } catch(Exception e) {};
				
				h.put(TABLE_FILETRANSFERMECHANISM1 + "_" + fieldName, Arrays.asList(fieldName,file,prodSourceFilePath,sitSourceFilePath,devSourceFilePath));
				h.put(TABLE_FILETRANSFERMECHANISM2 + "_" + fieldName, Arrays.asList(fileName,fileExt,controlFileName,controlFileExt));
				h.put(TABLE_SOURCEFEEDFILES + "_" + fieldName, Arrays.asList(fieldName,fileName,fileExt,fileType,"50",freq,"Pull",minNoFiles,freqSource,delimiter,integrationType));
				h.put(TABLE_FIRSTLOAD + "_" + fieldName, Arrays.asList(fieldName,fileName,fileExt,historicalLoad,"6 Months"));
				h.put(TABLE_FEED2TABLE + "_" + fieldName, Arrays.asList(fieldName,tableName,recordPerDay,fileType));
				h.put(TABLE_CONTROLFEED + "_" + fieldName, Arrays.asList(fieldName,fileName,fileExt,controlFileName,controlFileExt,controlFormat));
				h.put(TABLE_AVAILABILITY + "_" + fieldName, Arrays.asList(fieldName,fileName,availableTime));
			}
			return h;
		} finally {
		}
	}*/

	private static String getFieldValue(Row row,String column) {
		return row.getCell(CellReference.convertColStringToIndex(column)).toString();
	}

	private static String getFieldValue(String[] field, String column) {
		try { 
			return field[indexForColumn(column)].trim(); 
		} catch(Exception e) {
			return "";
		} 
	}

	private static int indexForColumn(String column) {
		return CellReference.convertColStringToIndex(column);
	}

	private static void printHashmap(HashMap<String,List> h) {
		h.entrySet().forEach(entry -> {
			System.out.println(entry.getKey() + " => " + entry.getValue());
		});
	}

	private static void printSortedHashmap(HashMap hmap) {
		Map<Integer, String> map = new TreeMap<Integer, String>(hmap); 
		Set set2 = map.entrySet();
		Iterator iterator2 = set2.iterator();
		while(iterator2.hasNext()) {
			Map.Entry me2 = (Map.Entry)iterator2.next();
			System.out.println(me2.getKey() + " => " + me2.getValue());
			System.out.println("=====================================");
		}
	}
}