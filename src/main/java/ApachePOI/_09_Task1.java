package ApachePOI;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class _09_Task1 {

    /**
     *  Çarpım tablosunu excele yazdırınız.
     *  1 x 1 = 1 şeklinde işaretleri de yazdırınız.
     *  sıfırdan excel oluşturarak.
     */

    public static void main(String[] args) throws IOException {

        String path="src/main/resources/CarpimTablosu.xlsx";

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Carpim");



//        Row row=sheet.createRow(0);
//        Cell cell=row.createCell(0);
//
//        cell.setCellValue("1");

        int rowCount=0;
        for(int i=1;i<=10;i++){
            for(int j=1;j<=10;j++){

                Row row=sheet.createRow(rowCount++); // her işlem için bir satır açar mı?

                Cell cell=row.createCell(0); cell.setCellValue(i);
                cell=row.createCell(1); cell.setCellValue("x");
                cell=row.createCell(2); cell.setCellValue(j);
                cell=row.createCell(3); cell.setCellValue(" =");
                cell=row.createCell(4); cell.setCellValue((i*j));

                //System.out.println(i+"x"+j+"="+(i*j));
            }
        }



        FileOutputStream outputStream=new FileOutputStream(path);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

    }
}