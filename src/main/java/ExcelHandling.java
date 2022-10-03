import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ExcelHandling {


  public static void main(String[] args) {
      String fileName = "D:/VEHICLE.xlsx";
      Workbook workbook = null;
      File file = new File(fileName);
      FileInputStream fileInputStream;


    try {

        fileInputStream = new FileInputStream(file);
        String fileExtension = fileName.substring(fileName.indexOf("."));
        System.out.println(fileExtension);
        if(fileExtension.equals(".xls")){
            workbook  = new HSSFWorkbook(new POIFSFileSystem(fileInputStream));
        }
        else if(fileExtension.equals(".xlsx")){
            workbook  = new XSSFWorkbook(fileInputStream);
        }
        else {
            System.out.println("Wrong File Type");
        }
        DataFormatter dataFormatter = new DataFormatter();
        Iterator<Sheet> sheets = workbook.sheetIterator();
        while (sheets.hasNext()) {
            Sheet sh = sheets.next();
            System.out.println("Sheet name is " + sh.getSheetName());
            System.out.println("----------------");
            Iterator<Row> iterator = sh.iterator();
            while (iterator.hasNext()) {
                Row row = iterator.next();
                Iterator<Cell> cellIterator = row.iterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = dataFormatter.formatCellValue(cell);
//                    if(cell.getCellType() == CellType.STRING){
//
//                    }
                    System.out.print(cellValue + "\t");
                }
                System.out.println();
            }
        }
    }
    catch (Exception e) {
        e.printStackTrace();
    }
  }
}
