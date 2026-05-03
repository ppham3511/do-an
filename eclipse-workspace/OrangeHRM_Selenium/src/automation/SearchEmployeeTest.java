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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class SearchEmployeeTest implements ITest {
    WebDriver driver;
    WebDriverWait wait;
    Workbook wb;
    Sheet sh;
    
    // Cấu hình cột
    int colExpected = 3;   // Cột D
    int colActualRes = 5;  // Cột F
    int colActualMsg = 6;  // Cột G
    int colStatus    = 7;  // Cột H

    int totalTC = 0;
    int matchedCount = 0;

    private ThreadLocal<String> testName = new ThreadLocal<>();

    @Override
    public String getTestName() {
        return testName.get();
    }

    @BeforeMethod
    public void setTestName(Method method, Object[] data) {
        // Cập nhật để hiển thị ID test case từ file Excel nếu cần, hoặc giữ nguyên method name
        testName.set(data[0].toString() + "_" + method.getName()); 
    }

    @BeforeClass
    public void setup() throws Exception {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Chọn File Excel Input cho Search Test");
        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
            throw new RuntimeException("Chưa chọn file đầu vào!");
        }

        File inputFile = chooser.getSelectedFile();
        FileInputStream fis = new FileInputStream(inputFile);
        wb = new XSSFWorkbook(fis);
        sh = wb.getSheetAt(0);

        // --- TẠO STYLE IN ĐẬM ---
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        // Tạo tiêu đề cột và áp dụng in đậm
        Row headerRow = sh.getRow(0);
        String[] headerTitles = {"Actual Result", "Actual Message", "Status (Match?)"};
        for (int i = 0; i < headerTitles.length; i++) {
            Cell cell = headerRow.createCell(colActualRes + i);
            cell.setCellValue(headerTitles[i]);
            cell.setCellStyle(headerStyle); // In đậm tiêu đề cột
        }

        // Setup trình duyệt
        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(op);
        driver.manage().window().maximize();
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        // Login một lần duy nhất
        driver.get("http://localhost/orangehrm/web/index.php/auth/login");
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("username"))).sendKeys("admin");
        driver.findElement(By.name("password")).sendKeys("Admin@1234");
        driver.findElement(By.xpath("//button[@type='submit']")).click();
        wait.until(ExpectedConditions.urlContains("dashboard"));
    }

    @DataProvider(name = "searchData")
    public Iterator<Object[]> getSearchData() {
        List<Object[]> data = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        for (int i = 1; i <= sh.getLastRowNum(); i++) {
            Row row = sh.getRow(i);
            if (row == null || df.formatCellValue(row.getCell(0)).isEmpty()) continue;
            
            data.add(new Object[]{
                df.formatCellValue(row.getCell(1)).trim(), // empName
                df.formatCellValue(row.getCell(2)).trim(), // empId
                df.formatCellValue(row.getCell(colExpected)).trim().toUpperCase(), // expRes
                i // rowIndex
            });
        }
        return data.iterator();
    }

    @Test(dataProvider = "searchData")
    public void testSearchEmployee(String empName, String empId, String expRes, int rowIndex) {
        totalTC++;
        String actRes = "FAIL";
        String actMsg = "";

        try {
            driver.get("http://localhost/orangehrm/web/index.php/pim/viewEmployeeList");
            wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//button[@type='submit']")));

            WebElement nameField = driver.findElement(By.xpath("//label[text()='Employee Name']/parent::div/following-sibling::div//input"));
            WebElement idField = driver.findElement(By.xpath("//label[text()='Employee Id']/parent::div/following-sibling::div/input"));
            
            nameField.sendKeys(Keys.CONTROL + "a", Keys.BACK_SPACE);
            idField.sendKeys(Keys.CONTROL + "a", Keys.BACK_SPACE);

            if (!empName.isEmpty()) nameField.sendKeys(empName);
            if (!empId.isEmpty()) idField.sendKeys(empId);

            driver.findElement(By.xpath("//button[@type='submit']")).click();
            Thread.sleep(2000); 

            List<WebElement> records = driver.findElements(By.className("oxd-table-card"));
            
            if (!records.isEmpty()) {
                actRes = "PASS";
                actMsg = "Tìm thấy " + records.size() + " bản ghi";
            } else {
                actRes = "FAIL";
                try {
                    WebElement infoMsg = driver.findElement(By.xpath("//span[contains(@class,'oxd-text') and (contains(.,'No Records Found') or contains(.,'Records Found'))]"));
                    actMsg = infoMsg.getText();
                } catch (Exception e) {
                    actMsg = "No Records Found";
                }
            }
        } catch (Exception e) {
            actRes = "ERROR";
            actMsg = e.getMessage().split("\n")[0];
        }

        Row row = sh.getRow(rowIndex);
        row.createCell(colActualRes).setCellValue(actRes);
        row.createCell(colActualMsg).setCellValue(actMsg);
        
        boolean isMatch = actRes.equalsIgnoreCase(expRes);
        if (isMatch) matchedCount++;
        row.createCell(colStatus).setCellValue(isMatch ? "PASS" : "FAIL");

        Assert.assertTrue(isMatch, "Dòng " + rowIndex + " không khớp mong đợi.");
    }

    @AfterClass
    public void tearDown() throws Exception {
        saveResultsWithReport();
        if (driver != null) driver.quit();
    }

    private void saveResultsWithReport() throws Exception {
        int summaryIdx = sh.getLastRowNum() + 2;
        double accuracy = (totalTC > 0) ? ((double) matchedCount / totalTC) * 100 : 0;

        // Tạo style in đậm cho phần báo cáo
        CellStyle boldStyle = wb.createCellStyle();
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        boldStyle.setFont(boldFont);

        String[] labels = {"Tổng số testcase:", "Số case khớp:", "Số case lệch:", "Độ chính xác:"};
        String[] values = {String.valueOf(totalTC), String.valueOf(matchedCount), String.valueOf(totalTC - matchedCount), String.format("%.2f%%", accuracy)};

        for (int j = 0; j < labels.length; j++) {
            Row r = sh.createRow(summaryIdx + j);
            Cell cellLabel = r.createCell(0);
            cellLabel.setCellValue(labels[j]);
            cellLabel.setCellStyle(boldStyle); // In đậm nhãn thống kê
            
            r.createCell(1).setCellValue(values[j]);
        }

        JFileChooser save = new JFileChooser();
        save.setDialogTitle("Lưu Báo Cáo Kết Quả Search");
        if (save.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            String path = save.getSelectedFile().getAbsolutePath();
            if (!path.endsWith(".xlsx")) path += ".xlsx";
            
            try (FileOutputStream fos = new FileOutputStream(path)) {
                wb.write(fos);
            }
            JOptionPane.showMessageDialog(null, "Lưu file thành công! Độ chính xác: " + String.format("%.2f%%", accuracy));
        }
        wb.close();
    }
}