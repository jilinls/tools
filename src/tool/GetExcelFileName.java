package tool;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.Reader;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetExcelFileName {

	public static void main(String[] args) throws Exception {
		
		// 获取路径下所有文件名
		String listPath = "C:\\";
		
		StringBuffer sb = new StringBuffer("");
		Reader reader = new FileReader(listPath + "list.txt");
		BufferedReader bufferedReader = new BufferedReader(reader);
		String string = null;

		FileInputStream fileIn = null;
		Workbook wb0 = null;
		Sheet sht0 = null;
		Row row0 = null;

		//共通インターフェースを扱える、WorkbookFactoryで読み込む
        Workbook wb = null;
        
		while ((string = bufferedReader.readLine()) != null) {
			
//			string = string.substring(
//					string.lastIndexOf("\\")+1,
//					string.lastIndexOf("."));
			
			wb = WorkbookFactory.create(new FileInputStream(string));
            //全セルを表示する
            for (Sheet sheet : wb ) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        System.out.print(cell.getRichStringCellValue());
                        System.out.print(" , ");
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
	            return cell.getCellFormula();
	
	        default:
	            return null;
	    }
	}

}
