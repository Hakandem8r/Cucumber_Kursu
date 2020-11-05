package ApachePOI;


import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class _07_WriteInTheExcel {
    public static void main(String[] args) throws IOException {

        String path="src/main/resources/WriteInTheExcelFile.xlsx";

        //dosya okuma işlemleri
        FileInputStream inputStream=new FileInputStream(path);
        Workbook workbook= WorkbookFactory.create(inputStream);
        Sheet sheet=workbook.getSheet("Sheet1");

        // hafızada yazma işlemleri
        Row row=sheet.createRow(0); // satır oluşturdum ama içinnd ehiç hücre yok
        Cell cell=row.createCell(0); // satırın içine bir hücre oluşturdum.

        cell.setCellValue("Merhaba dünya");
        inputStream.close(); // okuma işlemi kapatıldı.

        // yazma işlemi için exceli yazma modunda açıyorum.
        FileOutputStream outputStream=new FileOutputStream(path);
        workbook.write(outputStream);
        outputStream.close();
        // dosya kapatıldı.
    }
}
