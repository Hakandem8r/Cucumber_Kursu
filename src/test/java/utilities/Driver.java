package utilities;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeDriverService;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Driver {

    /**
     * Interview Sorusu
     * Singleton driver nedir?
     * Bütün projede aynı ve tek bir driverın kullanılmasıdır.
     */

    // Siz Bütün dersler için sınav kağıtları hazırlayan bir yöneticisiniz.
    // 1.Öğrenci geldi, sizden fizik dersi yazılı sorularını istedi. verdiniz.
    // 2.Öğrenci geldi, sizden kimya dersi yazılı sorularını istedi. verdiniz.
    // 3.Öğrenci geldi sizden yine Fizik yazılı soruları istedi, siz yeni soru hazırlamazsınız.
    //    aynı fizik sorularından vermelisiniz.
    // 4.Öğrenci geldi sizden kimya yazlı soruları istedi.Yeni soru hazılamazsını yine daha önce
    //    hazırladığınız Kimya sorularını verirsiniz.
    // Böylece ilk talep için daha önce yok ise yeni oluşturusunuz, ancak aynı talep tekrar gelirse
    // daha önce oluşturulmuşu vermeniz gerekir.
    // burda öğrencileri CLASS a yani Senaryolara benzetebiliriz. Sınav kağıtları ise DRİVER dır.

    // Her farklı browser için farklı bir static değişken havuzu oluşturduk.
    private static ThreadLocal<WebDriver> threadDriver=new ThreadLocal<>();
    public static ThreadLocal<String> threadBrowserName=new ThreadLocal<>();

    public static WebDriver getDriver()
    {
        // eğer browser adı gelmezse
        if (threadBrowserName.get()==null)
        {
            threadBrowserName.set("chrome");
        }

        if (threadDriver.get() == null)
        {
            switch (threadBrowserName.get())
            {
                case "firefox":
                    WebDriverManager.firefoxdriver().setup();
                    threadDriver.set( new FirefoxDriver() );
                    break;

                default:
                    WebDriverManager.chromedriver().setup();
                    threadDriver.set( new ChromeDriver() );
                    break;

            }
        }
        return threadDriver.get();
    }

    public static void quitDriver()
    {
        if (threadDriver.get() != null)
        {
            threadDriver.get().quit(); // ilgili driverdan quit edildi
            WebDriver driver=threadDriver.get(); // threadLocl deki ilgili driverının referansı alındı.
            driver=null;  // null değeri atanarak boşaltıldı.
            threadDriver.set(driver); // ve set edilerek ThreadLocal in içindeki saklanan null hale getirildi.
        }
    }



}