
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class App {

    public void writeIntoExcel(String file, String writeString) throws FileNotFoundException, IOException {

        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Sheet");

        Row row = sheet.createRow(0);

        Cell text = row.createCell(0);
        text.setCellValue(writeString);

        sheet.autoSizeColumn(0);

        book.write(new FileOutputStream(file));
        book.close();

    }

    public String readFromExcel(String file) throws IOException {
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Sheet");
        XSSFRow row = myExcelSheet.getRow(0);

        String name = row.getCell(0).getStringCellValue();

        myExcelBook.close();
        return name;
    }

}
