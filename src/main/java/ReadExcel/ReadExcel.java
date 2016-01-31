package ReadExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

/**
 * Created by gkoles on 2016.01.31..
 */
public class ReadExcel {
    public static void main(String[] args) {

        try {
            FileInputStream file = new FileInputStream(new File("demo.xlsx"));

            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheet("Company Data");

            //Iterate through rows one by one
            Iterator<Row> rowIterator = sheet.iterator();

            //Create map to store data in
            Map<String, Object[]> data = new TreeMap<String, Object[]>();


            while (rowIterator.hasNext()){

                Row row = rowIterator.next();

                //For each row, let's iterate through the cells
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()){

                    Cell nextCell = cellIterator.next();
                    int ColumnIndex = nextCell.getColumnIndex();

                    switch (ColumnIndex){

                        case 0:
                            if(nextCell.getCellType() == Cell.CELL_TYPE_STRING) System.out.println(nextCell.getStringCellValue());
                            else System.out.println((int) nextCell.getNumericCellValue());


                    }

                }


            }


        }
        catch (Exception e){

            e.printStackTrace();

        }

    }
}
