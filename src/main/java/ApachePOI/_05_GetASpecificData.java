package ApachePOI;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

/**
 * Verilen Exceldeki password bilgisini bir metoda aracılığı bularak yazdırınız,
 * yani metoda "password" kelimesi gönderilecek, dönen değr password un kendisi olacak.
 *
 */
public class _05_GetASpecificData {

    public static void main(String[] args) {

        System.out.print("Aranacak Kelime=");
        Scanner oku=new Scanner(System.in);
        String arananKelime=oku.nextLine();

        String donenKelime=bul(arananKelime);
        System.out.println(donenKelime);
    }

    static String bul(String aranacakKelime) {
        String donecekKelime="";

        String path="src/main/resources/LoginData.xlsx";
        Workbook workbook=null;
        try {
            FileInputStream inputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(inputStream);
        }
        catch(Exception ex)
        {
            ex.printStackTrace();
        }

        Sheet sheet= workbook.getSheet("Login");

        int rowCount= sheet.getPhysicalNumberOfRows();

        for(int i=0;i< rowCount ;i++)
        {
            //i.satırdaki 0 ın hücrenin değer arananan kelimeye eşit mi ?
            if (sheet.getRow(i).getCell(0).toString().equalsIgnoreCase(aranacakKelime))
            {
                //bulunan satırda kaç hücre var
                int cellCount=sheet.getRow(i).getPhysicalNumberOfCells();

                for(int j=0; j< cellCount;j++)
                {
                    donecekKelime+=sheet.getRow(i).getCell(j).toString()+" ";
                }
            }
        }
        return  donecekKelime;
    }

}