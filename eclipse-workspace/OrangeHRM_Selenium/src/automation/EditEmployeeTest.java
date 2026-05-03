package automation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.ITest;
import org.testng.annotations.*;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

public class EditEmployeeTest implements ITest {
    WebDriver driver;
    WebDriverWait wait;
    JavascriptExecutor js;
    Workbook wb;
    Sheet sh;

    int colActualRes = 9, colActualMsg = 10, colStatus = 11;
    int totalTC = 0, matchedCount = 0;
    private ThreadLocal<String> testName = new ThreadLocal<>();

    @Override public String getTestName() { return testName.get(); }
    @BeforeMethod public void setTestName(Method method, Object[] data) { testName.set(data[0].toString()); }

    @BeforeClass
    public void setup() throws Exception {
        JFileChooser chooser = new JFileChooser();
        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) throw new RuntimeException("Chưa chọn file!");
        wb = new XSSFWorkbook(new FileInputStream(chooser.getSelectedFile()));
        sh = wb.getSheetAt(0);

        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(op);
        driver.manage().window().maximize();
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        js = (JavascriptExecutor) driver;

        driver.get("http://localhost/orangehrm/web/index.php/auth/login");
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("username"))).sendKeys("admin");
        driver.findElement(By.name("password")).sendKeys("Admin@1234");
        driver.findElement(By.xpath("//button[@type='submit']")).click();
        wait.until(ExpectedConditions.urlContains("dashboard"));
    }

    @DataProvider(name = "editData")
    public Iterator<Object[]> getEditData() {
        List<Object[]> data = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        for (int i = 1; i <= sh.getLastRowNum(); i++) {
            Row row = sh.getRow(i);
            if (row == null || df.formatCellValue(row.getCell(0)).isEmpty()) continue;
            data.add(new Object[]{
                df.formatCellValue(row.getCell(0)).trim(),
                formatID(df.formatCellValue(row.getCell(1))),
                df.formatCellValue(row.getCell(3)).trim(),
                df.formatCellValue(row.getCell(4)).trim(),
                df.formatCellValue(row.getCell(5)).trim(),
                formatID(df.formatCellValue(row.getCell(6))),
                df.formatCellValue(row.getCell(7)).trim().toUpperCase(),
                i 
            });
        }
        return data.iterator();
    }

    @Test(dataProvider = "editData")
    public void testEditEmployee(String tcID, String targetID, String nFirst, String nMiddle, String nLast, 
                                 String nId, String expRes, int rowIndex) {
        totalTC++;
        String actRes = "FAIL";
        String actMsg = "";

        try {
            driver.get("http://localhost/orangehrm/web/index.php/pim/viewEmployeeList");
            WebElement sBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//label[text()='Employee Id']/parent::div/following-sibling::div/input")));
            forceInput(sBox, targetID);
            driver.findElement(By.xpath("//button[@type='submit']")).click();
            Thread.sleep(1500);

            if (driver.findElements(By.xpath("//span[text()='No Records Found']")).size() > 0) {
                actMsg = "Employee not found";
            } else {
                WebElement editBtn = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[.//i[contains(@class,'bi-pencil-fill')]]")));
                js.executeScript("arguments[0].click();", editBtn);
                
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("firstName")));
                
                // Ép thay đổi dữ liệu bằng hàm forceInput mới
                forceInput(driver.findElement(By.name("firstName")), nFirst);
                forceInput(driver.findElement(By.name("middleName")), nMiddle);
                forceInput(driver.findElement(By.name("lastName")), nLast);
                
                WebElement idInput = driver.findElement(By.xpath("//label[text()='Employee Id']/parent::div/following-sibling::div/input"));
                forceInput(idInput, nId);

                // Cuộn tới nút Save và Click bằng JS để tránh bị che khuất
                WebElement saveBtn = driver.findElement(By.xpath("(//button[@type='submit'])[1]"));
                js.executeScript("arguments[0].scrollIntoView(true);", saveBtn);
                Thread.sleep(500);
                js.executeScript("arguments[0].click();", saveBtn);
                
                Thread.sleep(3000); // Đợi lưu hoàn tất

                // KIỂM TRA KẾT QUẢ
                List<WebElement> errors = driver.findElements(By.cssSelector(".oxd-input-group__message, .oxd-input-field-error-message, .--error"));
                
                if (!errors.isEmpty()) {
                    actRes = "FAIL";
                    actMsg = errors.stream().map(WebElement::getText).filter(t->!t.isEmpty()).distinct().collect(Collectors.joining("; "));
                } else {
                    try {
                        WebElement toast = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(".oxd-toast-content")));
                        actMsg = toast.getText().trim();
                        actRes = actMsg.toLowerCase().contains("success") ? "PASS" : "FAIL";
                    } catch (Exception e) {
                        actRes = "FAIL";
                        actMsg = "No Success Toast - Data might not be saved";
                    }
                }
            }
        } catch (Exception e) {
            actRes = "FAIL";
            actMsg = "Error: " + e.getMessage();
        }

        Row row = sh.getRow(rowIndex);
        row.createCell(colActualRes).setCellValue(actRes);
        row.createCell(colActualMsg).setCellValue(actMsg);
        boolean isMatch = actRes.equalsIgnoreCase(expRes);
        if (isMatch) matchedCount++;
        row.createCell(colStatus).setCellValue(isMatch ? "PASS" : "FAIL");
        Assert.assertTrue(isMatch, "TC " + tcID + " không khớp. Thực tế: " + actMsg);
    }

    // HÀM QUAN TRỌNG: Đảm bảo dữ liệu được thay đổi thực sự
    private void forceInput(WebElement el, String val) {
        // 1. Xóa bằng phím tắt
        el.sendKeys(Keys.CONTROL + "a");
        el.sendKeys(Keys.BACK_SPACE);
        
        // 2. Xóa bằng Javascript (để chắc chắn trống hoàn toàn)
        js.executeScript("arguments[0].value = '';", el);
        
        // 3. Nhập giá trị mới
        if (val != null && !val.isEmpty()) {
            el.sendKeys(val);
        }
        
        // 4. Kích hoạt toàn bộ sự kiện để trang web nhận diện có sự thay đổi
        js.executeScript("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", el);
        js.executeScript("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", el);
        js.executeScript("arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));", el);
    }

    private String formatID(String val) {
        if (val == null || val.isEmpty()) return "";
        if (val.endsWith(".0")) return val.substring(0, val.length() - 2);
        return val.trim();
    }

    @AfterClass
    public void tearDown() throws Exception {
        saveResults();
        if (driver != null) driver.quit();
    }

    private void saveResults() throws Exception {
        double acc = (totalTC > 0) ? ((double) matchedCount / totalTC) * 100 : 0;
        Row r = sh.createRow(sh.getLastRowNum() + 2);
        r.createCell(0).setCellValue("Accuracy: " + String.format("%.1f%%", acc));

        JFileChooser save = new JFileChooser();
        if (save.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            FileOutputStream fos = new FileOutputStream(save.getSelectedFile().getAbsolutePath() + ".xlsx");
            wb.write(fos);
            fos.close();
        }
        wb.close();
    }
}