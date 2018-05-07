/**
 * lendo arquivo .xlsx com APACHE POI
 */

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class FileUtil {

    public static void main(String[] args) throws IOException {
        File myFile = new File("C://exames.xlsx");
        FileInputStream fis = new FileInputStream(myFile);

        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
        Iterator<Row> rowIterator = mySheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                int z = cell.getColumnIndex();

                switch (z) {
                    case 0:
                        System.out.print(cell.toString() + "\t");
                        break;
                    case 1:
                        System.out.print(cell.toString() + "\t");
                        break;
                    case 2:
                        System.out.print(cell.toString() + "\t");
                        break;
                    default:
                }
            }

            System.out.println(" ");
        }
    }
}
