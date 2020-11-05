package ApachePOI;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class _03_ApachePOIStart {
    public static void main(String[] args) throws IOException {

        String path = "src/main/resources/ApacheExcel1.xlsx";

        //dosyayı yani Excel'i okumaya açtım.
        FileInputStream dosyaOkumaYolu = new FileInputStream(path);

        //Bunun üzerinden Çalışma kitabını alıyorum.
        Workbook calismaKitabi = WorkbookFactory.create(dosyaOkumaYolu);

        //İstediğim isimdeki çalışma sayfasını alıyorum
        Sheet calismaSayfasi = calismaKitabi.getSheetAt(0);

        //İstenen Satırı alıyorum
        Row satir = calismaSayfasi.getRow(0);

        //İstenen satırdaki işlenen hücre alınıyor
        Cell hucre = satir.getCell(0);

        System.out.println(hucre);
    }
}
