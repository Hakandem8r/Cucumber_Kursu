package ApachePOI;

import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class _04_ApachePOIGetAllData {
    public static void main(String[] args) throws IOException {

        String path="src/main/resources/ApacheExcel1.xlsx";
        FileInputStream inputStream= new FileInputStream(path);
        Workbook workbook= WorkbookFactory.create(inputStream);
        Sheet sheet= workbook.getSheet("Sheet1");

        // calisma sayfasındaki toplam satır sayısını veriyor.
        int rowCount= sheet.getPhysicalNumberOfRows();
        System.out.println("Satir Sayısı="+rowCount);

        for(int i=0;i<rowCount ;i++)//satır sayısı kadar dönecek
        {
            Row row=sheet.getRow(i); // i.Satır alındı
            int cellCount = row.getPhysicalNumberOfCells(); // bu satırdaki toplam hücre sayısı alındı.

            for(int j=0;j< cellCount;j++ )//i.satırdaki hücre sayısı kadar dönecek
            {
                Cell cell = row.getCell(j); //bu satırdaki siradaki hücreyi aldık
                // System.out.print(cell+" ");
                System.out.printf("%-15s",cell);

                /*
                Cell cell=row.getCell(j);     //bu satırdaki sıradaki hucreyi aldım.
                //System.out.print(cell+ " ");

              //  System.out.printf("%10s",cell); // saga dayalı 10 haneli String yazdırdı.
              //  System.out.printf("%15s",cell); // saga dayalı 15 haneli String yazdırdı.BÖYLECE satırlar ayrıştılar.
                System.out.printf("%-15s",cell);  // Sola dayalı 15 haneli String yazdırdı
        // prıntlerin hepsını ayrı ayrı aç farklarını farket.
                 */
            }
            System.out.println();
        }
    }
}