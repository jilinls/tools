package tool;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.Reader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class GetExcelFileName {

	public static void main(String[] args) throws Exception {
		
		String listPath = "C:\\";
		StringBuffer sb = new StringBuffer("");
		Reader reader = new FileReader(listPath + "list.txt");
		BufferedReader bufferedReader = new BufferedReader(reader);
		String string = null;
		//共通インターフェースを扱える、WorkbookFactoryで読み込む
        Workbook wb = null;
        String str = "";
        File file = null;
        
		while ((string = bufferedReader.readLine()) != null) {
			
//			string = string.substring(
//					string.lastIndexOf("\\")+1,
//					string.lastIndexOf("."));
			
			file = new File(string);
			if (!file.exists()) {
				System.out.println(string + "不存在");
				continue;
			}
			
			wb = WorkbookFactory.create(file);
			
            //全セルを表示する
            for (Sheet sheet : wb ) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                    	str = getCellValue(cell).toString();
                    	if (str.contains("\n") || str.contains("\r\n")) {
                    		str = str.replace("\n", ",");
                    	}
                        System.out.println(str);
                    }
                    System.out.println();
                }
            }
			sb.append(string);
			System.out.println(string);
		}
		bufferedReader.close();
		reader.close();
	}

	private static Object getCellValue(Cell cell) {
	        switch (cell.getCellType()) {
	        case Cell.CELL_TYPE_STRING:
	            return cell.getRichStringCellValue().getString();
	        case Cell.CELL_TYPE_NUMERIC:
	            if (DateUtil.isCellDateFormatted(cell)) {
	                return cell.getDateCellValue();
	            } else {
	                return cell.getNumericCellValue();
	            }
	        case Cell.CELL_TYPE_BOOLEAN:
	            return cell.getBooleanCellValue();
	        case Cell.CELL_TYPE_FORMULA:
	        	// return cell.getCellFormula();
	            return cell.getRichStringCellValue();
	        default:
	            return null;
	    }
	}
}
