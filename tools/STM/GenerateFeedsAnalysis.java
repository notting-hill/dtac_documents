import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.commons.io.FilenameUtils;

public class GenerateFeedsAnalysis {
	public static void main(String[] args) throws Exception {
		generate();
	}

	public static void generate() throws Exception {
		File outFile = new File("feedsAnalysis.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		File file = new File("tables.txt");
		BufferedReader br = null; 
		try {
			br = new BufferedReader(new FileReader(file)); 
			String line;
			HashMap h = new HashMap();
			while((line = br.readLine()) != null){
				XSSFSheet sheetSummary = workbook.createSheet(line.trim().toLowerCase());
			}
		} catch(Exception e) {
		}
		try (FileOutputStream outputStream = new FileOutputStream(outFile)) {
			workbook.write(outputStream);
		}
	}
}