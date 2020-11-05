package ApachePOI;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class _08_NewExcelWrite {
    public static void main(String[] args) throws IOException {

        // yeni workbook oluştur, sonra sheet oluştur, sonra row oluştur, sonra cell oluştur.
        XSSFWorkbook workbook=new XSSFWorkbook(); // hazfızada workbook oluştur.
        XSSFSheet sheet=workbook.createSheet("Calisma Sayfasi"); // hazıfazda sheet oluştur.

        Row row=sheet.createRow(0); // satır oluştur
        Cell cell=row.createCell(0); // hucre oluştur

        cell.setCellValue("Merhaba Dünya biz geldik."); // hucreye hafızada yaz.

        // dosyayı oluşturduk.
        FileOutputStream outputStream=new FileOutputStream("src/main/resources/NewExcel.xlsx");
        workbook.write(outputStream); // içine hafızadaki bilgileri yazdık.
        workbook.close(); // hafızayı boşalttık
        outputStream.close(); // dosyayı kapattık.
    }

}
