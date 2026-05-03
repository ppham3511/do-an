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

public class ApplyLeaveTest implements ITest {
    WebDriver driver;
    WebDriverWait wait;
    JavascriptExecutor js;
    Workbook wb;
    Sheet sh;
    
    // Vị trí các cột kết quả trong Excel
    int colActualRes = 8;  
    int colActualMsg = 9; 
    int colStatus    = 10;

    int totalTC = 0;
    int matchedCount = 0;
    private ThreadLocal<String> testName = new ThreadLocal<>();

    @Override
    public String getTestName() { return testName.get(); }

    @BeforeMethod
    public void setTestName(Method method, Object[] data) {
        testName.set(data[0].toString()); 
    }

    @BeforeClass
    public void setup() throws Exception {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Chọn File Excel Input cho Apply Leave");
        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
            throw new RuntimeException("Chưa chọn file input!");
        }

        FileInputStream fis = new FileInputStream(chooser.getSelectedFile());
        wb = new XSSFWorkbook(fis);
        sh = wb.getSheetAt(0);

        // --- TẠO STYLE IN ĐẬM ---
        CellStyle headerStyle = wb.createCellStyle();
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        headerStyle.setFont(boldFont);

        // Khởi tạo header cho kết quả và áp dụng in đậm
        Row header = sh.getRow(0);
        String[] titles = {"Actual Result", "Actual Message", "Status Match"};
        for (int i = 0; i < titles.length; i++) {
            Cell c = header.createCell(colActualRes + i);
            c.setCellValue(titles[i]);
            c.setCellStyle(headerStyle); 
        }

        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(op);
        driver.manage().window().maximize();
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        js = (JavascriptExecutor) driver;

        driver.get("http://localhost/orangehrm/web/index.php/auth/login");
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("username"))).sendKeys("cuongpv");
        driver.findElement(By.name("password")).sendKeys("Cuong@1234");
        driver.findElement(By.xpath("//button[@type='submit']")).click();
        wait.until(ExpectedConditions.urlContains("dashboard"));
    }

    @DataProvider(name = "leaveData")
    public Iterator<Object[]> getLeaveData() {
        List<Object[]> data = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        for (int i = 1; i <= sh.getLastRowNum(); i++) {
            Row row = sh.getRow(i);
            if (row == null || df.formatCellValue(row.getCell(0)).isEmpty()) continue;
            data.add(new Object[]{
                df.formatCellValue(row.getCell(0)).trim(),
                df.formatCellValue(row.getCell(1)).trim(),
                df.formatCellValue(row.getCell(2)).trim(),
                df.formatCellValue(row.getCell(3)).trim(),
                df.formatCellValue(row.getCell(4)).trim(),
                df.formatCellValue(row.getCell(5)).trim(),
                df.formatCellValue(row.getCell(6)).trim().toUpperCase(),
                i 
            });
        }
        return data.iterator();
    }

    @Test(dataProvider = "leaveData")
    public void testApplyLeave(String tcID, String type, String from, String to, 
                               String partial, String duration, String expRes, int rowIndex) {
        totalTC++;
        String actRes = "FAIL";
        String actMsg = "";

        try {
            driver.get("http://localhost/orangehrm/web/index.php/leave/applyLeave");
            Thread.sleep(2000);

            if (!type.isEmpty()) selectDropdown("Leave Type", type);
            if (!from.isEmpty()) fillInput("From Date", from);
            if (!to.isEmpty()) fillInput("To Date", to);
            
            Thread.sleep(1000);
            if (!partial.isEmpty()) selectDropdown("Partial Days", partial);
            if (!duration.isEmpty()) selectDropdown("Duration", duration);

            WebElement applyBtn = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@type='submit']")));
            js.executeScript("arguments[0].click();", applyBtn);
            
            Thread.sleep(2500);

            List<WebElement> dialogs = driver.findElements(By.xpath("//div[contains(@class, 'orangehrm-dialog-popup')]"));
            List<WebElement> errors = driver.findElements(By.cssSelector(".oxd-input-group__message"));

            if (!dialogs.isEmpty()) {
                actRes = "FAIL";
                actMsg = dialogs.get(0).getText().replace("\n", " ");
                try { driver.findElement(By.xpath("//button[contains(.,'Ok')]")).click(); } catch (Exception e) {}
            } 
            else if (!errors.isEmpty()) {
                actRes = "FAIL";
                actMsg = errors.stream()
                               .map(WebElement::getText)
                               .map(String::trim)
                               .filter(t -> !t.isEmpty())
                               .distinct()
                               .collect(Collectors.joining("; "));
            } 
            else {
                try {
                    WebElement toast = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(".oxd-toast-content")));
                    actMsg = toast.getText().replace("×", "").trim();
                    actRes = actMsg.toLowerCase().contains("success") ? "PASS" : "FAIL";
                } catch (Exception e) {
                    actRes = "FAIL";
                    actMsg = "No message captured";
                }
            }
        } catch (Exception e) {
            actRes = "FAIL";
            actMsg = "System Error: " + e.getMessage().split("\n")[0];
        }

        Row row = sh.getRow(rowIndex);
        row.createCell(colActualRes).setCellValue(actRes);
        row.createCell(colActualMsg).setCellValue(actMsg);
        
        boolean isMatch = actRes.equalsIgnoreCase(expRes);
        if (isMatch) matchedCount++;
        row.createCell(colStatus).setCellValue(isMatch ? "PASS" : "FAIL");

        Assert.assertTrue(isMatch, "TC " + tcID + " kết quả không khớp. Thực tế: " + actRes);
    }

    private void selectDropdown(String label, String val) {
        try {
            WebElement dd = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[text()='" + label + "']/parent::div/following-sibling::div//div[@class='oxd-select-wrapper']")));
            dd.click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@role='listbox']//span[text()='" + val + "']"))).click();
        } catch (Exception e) {}
    }

    private void fillInput(String label, String val) {
        try {
            WebElement el = driver.findElement(By.xpath("//label[text()='" + label + "']/parent::div/following-sibling::div//input"));
            el.sendKeys(Keys.CONTROL + "a", Keys.BACK_SPACE);
            el.sendKeys(val);
            js.executeScript("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", el);
        } catch (Exception e) {}
    }

    @AfterClass
    public void tearDown() throws Exception {
        saveResults();
        if (driver != null) driver.quit();
    }

    private void saveResults() throws Exception {
        double acc = (totalTC > 0) ? ((double) matchedCount / totalTC) * 100 : 0;
        int start = sh.getLastRowNum() + 2;

        // --- TẠO STYLE IN ĐẬM CHO TỔNG KẾT ---
        CellStyle boldStyle = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        boldStyle.setFont(font);

        // Bổ sung dòng "Số case lệch"
        String[] labels = {"Tổng số TC:", "Số case khớp (PASS):", "Số case lệch (FAIL):", "Độ chính xác:"};
        Object[] values = {totalTC, matchedCount, (totalTC - matchedCount), String.format("%.2f%%", acc)};

        for (int i = 0; i < labels.length; i++) {
            Row r = sh.createRow(start + i);
            Cell cLabel = r.createCell(0);
            cLabel.setCellValue(labels[i]);
            cLabel.setCellStyle(boldStyle); 
            
            r.createCell(1).setCellValue(values[i].toString());
        }

        JFileChooser save = new JFileChooser();
        save.setDialogTitle("Lưu file kết quả Apply Leave");
        if (save.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            String path = save.getSelectedFile().getAbsolutePath();
            if (!path.endsWith(".xlsx")) path += ".xlsx";
            try (FileOutputStream fos = new FileOutputStream(path)) {
                wb.write(fos);
            }
            JOptionPane.showMessageDialog(null, "Xuất báo cáo thành công! Độ chính xác: " + String.format("%.2f%%", acc));
        }
        wb.close();
    }
}