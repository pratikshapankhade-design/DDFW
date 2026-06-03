import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadDataFromExcel {
    public static void main(String[] args) throws IOException {

        FileInputStream file= new FileInputStream("C:\\Users\\272197\\IdeaProjects\\DDFW\\TestData\\Data.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(file);
        XSSFSheet sheet = wb.getSheet("Sheet1");
        int totalRows= sheet.getLastRowNum();
        int totalCell = sheet.getRow(1).getLastCellNum();
        System.out.println("Number of rows:" +totalRows);
        System.out.println("Number of Cells:" +totalCell);
        for (int r=0;r<=totalRows;r++)
        {
           XSSFRow currentRow = sheet.getRow(r);
            for(int c=0;c<totalCell;c++)
            {
                XSSFCell cell= currentRow.getCell(c) ;
               System.out.println(cell.toString());
            }
        }
        wb.close();
        file.close();
    }

}
