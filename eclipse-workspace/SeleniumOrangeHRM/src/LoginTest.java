
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class LoginTest {
    public static void main(String[] args) {

        // Set đường dẫn ChromeDriver
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");

        WebDriver driver = new ChromeDriver();

        // Mở website
        driver.get("https://opensource-demo.orangehrmlive.com/");
        driver.manage().window().maximize();

        // Nhập username
        driver.findElement(By.name("username")).sendKeys("admin");

        // Nhập password
        driver.findElement(By.name("password")).sendKeys("Admin@1234");

        // Click login
        driver.findElement(By.xpath("//button[@type='submit']")).click();

        // Kiểm tra login thành công (title hoặc URL)
        String currentUrl = driver.getCurrentUrl();
        if(currentUrl.contains("dashboard")) {
            System.out.println("Login SUCCESS");
        } else {
            System.out.println("Login FAIL");
        }

        // Đóng trình duyệt
        driver.quit();
    }
}