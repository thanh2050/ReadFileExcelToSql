import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;



public class InsertIntoMySQL {

    public static void main(String[] args) throws Exception {

        try {

            Class forName = Class.forName("com.mysql.jdbc.Driver");
            Connection con = null;
            con = DriverManager.getConnection("jdbc:mysql://localhost:3306/test1", "root", "thanhlan20");
            con.setAutoCommit(false);
            PreparedStatement pstm = null;
            FileInputStream input = new FileInputStream("D:\\Web\\Employee.xls");
            POIFSFileSystem fs = new POIFSFileSystem(input);
            Workbook workbook;
            workbook = WorkbookFactory.create(fs);
            Sheet sheet = workbook.getSheetAt(0);
            Row row;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                row = (Row) sheet.getRow(i);
                String name = row.getCell(0).getStringCellValue();
                String add = row.getCell(1).getStringCellValue();

                int  contact = (int) row.getCell(2).getNumericCellValue();

                String email = row.getCell(3).getStringCellValue();

                String sql = "INSERT INTO employee (name, address, contactNo, email) VALUES('" + name + "','" + add + "'," + contact + ",'" + email + "')";
                pstm = (PreparedStatement) con.prepareStatement(sql);
                pstm.execute();
                System.out.println("Import rows " + i);
            }
            con.commit();
            pstm.close();
            con.close();
            input.close();
            System.out.println("Success import excel to mysql table");
        } catch (IOException e) {
        }
    }
}