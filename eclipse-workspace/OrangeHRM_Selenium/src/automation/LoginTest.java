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

public class LoginTest implements ITest {
    WebDriver driver;
    WebDriverWait wait;
    Workbook wb;
    Sheet sh;
    int lastCol;
    
    int totalTestCases = 0;
    int matched = 0;

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
        chooser.setDialogTitle("Chọn File Excel Input cho Login");
        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
            throw new RuntimeException("Chưa chọn file!");
        }

        FileInputStream fis = new FileInputStream(chooser.getSelectedFile());
        wb = new XSSFWorkbook(fis);
        sh = wb.getSheetAt(0);

        // --- TẠO STYLE IN ĐẬM ---
        CellStyle headerStyle = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        headerStyle.setFont(font);

        // Khởi tạo header cho kết quả và áp dụng in đậm
        Row header = sh.getRow(0);
        lastCol = header.getLastCellNum();
        String[] resultHeaders = {"Actual Result", "Actual Message", "Status Match"};
        for (int i = 0; i < resultHeaders.length; i++) {
            Cell cell = header.createCell(lastCol + i);
            cell.setCellValue(resultHeaders[i]);
            cell.setCellStyle(headerStyle); // Áp dụng in đậm tiêu đề cột
        }

        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(op);
        driver.manage().window().maximize();
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
    }

    @DataProvider(name = "loginData")
    public Iterator<Object[]> getLoginData() {
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
                i
            });
        }
        return data.iterator();
    }

    @Test(dataProvider = "loginData")
    public void testLogin(String tcID, String user, String pass, String expRes, String expMsg, int rowIndex) {
        totalTestCases++;
        String actualRes = "FAIL";
        String actualMsg = "";

        try {
            driver.manage().deleteAllCookies();
            driver.get("http://localhost/orangehrm/web/index.php/auth/login");

            WebElement txtUser = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("username")));
            WebElement txtPass = driver.findElement(By.name("password"));
            WebElement btnLogin = driver.findElement(By.xpath("//button[@type='submit']"));

            txtUser.sendKeys(user);
            txtPass.sendKeys(pass);
            btnLogin.click();

            Thread.sleep(2000); 

            if (driver.getCurrentUrl().contains("dashboard")) {
                actualRes = "PASS";
                actualMsg = "Dashboard";
            } else {
                List<WebElement> errors = driver.findElements(By.xpath(
                    "//span[contains(@class,'oxd-input-field-error-message')] | //div[@role='alert']"
                ));

                if (!errors.isEmpty()) {
                    actualMsg = errors.stream()
                                      .map(WebElement::getText)
                                      .map(String::trim)
                                      .filter(t -> !t.isEmpty())
                                      .distinct()
                                      .collect(Collectors.joining(", "));
                } else {
                    actualMsg = "Error message not found";
                }
            }
        } catch (Exception e) {
            actualMsg = "Exception: " + e.getMessage().split("\n")[0];
        }

        Row row = sh.getRow(rowIndex);
        row.createCell(lastCol).setCellValue(actualRes);
        row.createCell(lastCol + 1).setCellValue(actualMsg);
        
        boolean isMatch = actualRes.equalsIgnoreCase(expRes);
        if (isMatch) matched++;
        row.createCell(lastCol + 2).setCellValue(isMatch ? "PASS" : "FAIL");

        Assert.assertTrue(isMatch, "TC " + tcID + " không khớp mong đợi.");
    }

    @AfterClass
    public void tearDown() throws Exception {
        saveResultsWithReport();
        if (driver != null) driver.quit();
    }

    private void saveResultsWithReport() throws Exception {
        double accuracy = (totalTestCases > 0) ? ((double) matched / totalTestCases) * 100 : 0;
        int summaryStartRow = sh.getLastRowNum() + 2;

        // Tạo style in đậm cho phần tổng kết
        CellStyle boldStyle = wb.createCellStyle();
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        boldStyle.setFont(boldFont);

        String[] labels = { "Tổng số Test Case:", "Số case khớp (PASS):", "Số case lệch (FAIL):", "Độ chính xác:"};
        Object[] values = { totalTestCases, matched, (totalTestCases - matched), String.format("%.2f%%", accuracy)};

        for (int i = 0; i < labels.length; i++) {
            Row r = sh.createRow(summaryStartRow + i);
            Cell cLabel = r.createCell(0);
            cLabel.setCellValue(labels[i]);
            cLabel.setCellStyle(boldStyle); // In đậm nhãn bên trái
            
            r.createCell(1).setCellValue(values[i].toString());
        }

        JFileChooser save = new JFileChooser();
        save.setDialogTitle("Lưu Báo Cáo Kết Quả Login");
        if (save.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            String path = save.getSelectedFile().getAbsolutePath();
            if(!path.endsWith(".xlsx")) path += ".xlsx";
            
            try (FileOutputStream fos = new FileOutputStream(path)) {
                wb.write(fos);
            }
            JOptionPane.showMessageDialog(null, "Xuất báo cáo thành công!\nĐộ chính xác: " + accuracy + "%");
        }
        wb.close();
    }
}