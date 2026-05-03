package automation;

import java.time.Duration;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class LoginPage {
    private WebDriver driver;
    private WebDriverWait wait;

    private By userField = By.name("username");
    private By passField = By.name("password");
    private By loginBtn = By.xpath("//button[@type='submit']");
    private By alertMsg = By.xpath("//p[contains(@class, 'oxd-alert-content-text')]");
    // Locator tìm tất cả các thông báo "Required"
    private By requiredMsgs = By.xpath("//span[contains(@class, 'oxd-input-field-error-message')]");

    public LoginPage(WebDriver driver) {
        this.driver = driver;
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(10));
    }

    public void login(String user, String pass) {
        // Chờ ô nhập liệu sẵn sàng
        wait.until(ExpectedConditions.visibilityOfElementLocated(userField)).sendKeys(user);
        driver.findElement(passField).sendKeys(pass);
        driver.findElement(loginBtn).click();
    }

    public String getInvalidCredentialsMessage() {
        return wait.until(ExpectedConditions.visibilityOfElementLocated(alertMsg)).getText();
    }

    public int getRequiredMessagesCount() {
        // Đợi ít nhất 1 thông báo lỗi xuất hiện trước khi đếm
        try {
            wait.until(ExpectedConditions.presenceOfElementLocated(requiredMsgs));
            List<WebElement> errors = driver.findElements(requiredMsgs);
            return errors.size();
        } catch (Exception e) {
            return 0;
        }
    }

    public boolean isDashboardDisplayed() {
        try {
            return wait.until(ExpectedConditions.urlContains("dashboard"));
        } catch (Exception e) {
            return false;
        }
    }
}