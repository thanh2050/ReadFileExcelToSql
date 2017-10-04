import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.io.FileOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcelDemo {

    public static void main(String[] args) throws IOException {

        // Đọc một file XSL.
        FileInputStream inputStream = new FileInputStream(new File("D:/Web/Employee.xls"));

        // Đối tượng workbook cho file XSL.
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        // Lấy ra sheet đầu tiên từ workbook
        HSSFSheet sheet = workbook.getSheetAt(0);

        // Lấy ra Iterator cho tất cả các dòng của sheet hiện tại.
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Lấy Iterator cho tất cả các cell của dòng hiện tại.
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch(cell.getCellType()) {
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + "\t\t");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t\t");
                        break;
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t\t");
                        break;
                }
            }
            System.out.println("");
        }
        inputStream.close();
        FileOutputStream out =
                new FileOutputStream(new File("D:/Web/Employee1.xls"));
        workbook.write(out);
        out.close();
        }
    }
