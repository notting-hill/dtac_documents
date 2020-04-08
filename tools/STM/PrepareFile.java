import java.io.*;
import java.util.*;

public class PrepareFile {

	public static void main(String[] args) {
		process();
	}

	private static void process() {
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

}