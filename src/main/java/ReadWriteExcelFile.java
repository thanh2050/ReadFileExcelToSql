import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;


    public class ReadWriteExcelFile {

        public static void readXLSFile() throws IOException
        {
            InputStream ExcelFileToRead = new FileInputStream("D:/Web/Employee.xls");
            HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
            HSSFSheet sheet=wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;

            Iterator rows = sheet.rowIterator();
            while (rows.hasNext())
            {
                row=(HSSFRow) rows.next();
                Iterator cells = row.cellIterator();
                while (cells.hasNext())
                {
                    cell=(HSSFCell) cells.next();
                    if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING)
                    {
                        System.out.print(cell.getStringCellValue()+" ");
                    }
                    else if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC)
                    {
                        System.out.print(cell.getNumericCellValue()+" ");
                    }

                }
                System.out.println();
            }

        }

        public static void writeXLSFile() throws IOException {

            String excelFileName = "D:/Web/Employee1.xls";

            String sheetName = "Sheet1";

            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet(sheetName) ;

            //iterating r number of rows
            for (int r=0;r < 5; r++ )
            {
                HSSFRow row = sheet.createRow(r);

                //iterating c number of columns
                for (int c=0;c < 5; c++ )
                {
                    HSSFCell cell = row.createCell(c);
                    cell.setCellValue("Cell "+r+" "+c);
                }
            }

            FileOutputStream fileOut = new FileOutputStream(excelFileName);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
        }

        public static void main(String[] args) throws IOException {

            writeXLSFile();
            readXLSFile();

            }

    }
