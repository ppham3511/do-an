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

public class AddEmployeeTest implements ITest {
    WebDriver driver;
    WebDriverWait wait;
    Workbook wb;
    Sheet sh;
    int colActualRes = 9;  
    int colActualMsg = 10; 
    int colStatus    = 11;

    // Biến thống kê
    int totalTC = 0;
    int matchedCount = 0;

    // Biến để ẩn tham số trên giao diện TestNG
    private ThreadLocal<String> testName = new ThreadLocal<>();

    @Override
    public String getTestName() {
        return testName.get();
    }

    @BeforeMethod
    public void setTestName(Method method) {
        testName.set(method.getName()); // Chỉ hiện "testAddEmployee" trên giao diện TestNG
    }

    @BeforeClass
    public void setup() throws Exception {
        // 1. CHỌN FILE EXCEL
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Chọn File Excel Input");
        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
            throw new RuntimeException("Chưa chọn file input!");
        }

        File inputFile = chooser.getSelectedFile();
        FileInputStream fis = new FileInputStream(inputFile);
        wb = new XSSFWorkbook(fis);
        sh = wb.getSheetAt(0);

        // 2. THIẾT LẬP TIÊU ĐỀ CỘT KẾT QUẢ
        Row headerRow = sh.getRow(0);
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        String[] headers = {"Actual Result", "Actual Message", "Status (Match?)"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(colActualRes + i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // 3. SETUP SELENIUM & LOGIN MỘT LẦN
        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(op);
        driver.manage().window().maximize();
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        driver.get("http://localhost/orangehrm/web/index.php/auth/login");
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("username"))).sendKeys("admin");
driver.findElement(By.name("password")).sendKeys("Admin@1234");
        driver.findElement(By.xpath("//button[@type='submit']")).click();
        wait.until(ExpectedConditions.urlContains("dashboard"));
    }

    @DataProvider(name = "employeeData")
    public Iterator<Object[]> getEmployeeData() {
        List<Object[]> data = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        for (int i = 1; i <= sh.getLastRowNum(); i++) {
            Row row = sh.getRow(i);
            if (row == null || df.formatCellValue(row.getCell(0)).isEmpty()) continue;
            
            // Đọc dữ liệu từng dòng
            data.add(new Object[]{
                df.formatCellValue(row.getCell(0)).trim(), // tcID
                df.formatCellValue(row.getCell(1)).trim(), // fName
                df.formatCellValue(row.getCell(2)).trim(), // mName
                df.formatCellValue(row.getCell(3)).trim(), // lName
                df.formatCellValue(row.getCell(4)).trim(), // eId
                df.formatCellValue(row.getCell(5)).trim(), // photo
                df.formatCellValue(row.getCell(6)).trim(), // expRes
                i // rowIndex
            });
        }
        return data.iterator();
    }

    @Test(dataProvider = "employeeData")
    public void testAddEmployee(String tcID, String fName, String mName, String lName, 
                                String eId, String photo, String expRes, int rowIndex) {
        totalTC++;
        String actRes = "FAIL";
        String actMsg = "";

        try {
            driver.get("http://localhost/orangehrm/web/index.php/pim/addEmployee");
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("firstName"))).sendKeys(fName);
            driver.findElement(By.name("middleName")).sendKeys(mName);
            driver.findElement(By.name("lastName")).sendKeys(lName);
            
            WebElement idInp = driver.findElement(By.xpath("//label[text()='Employee Id']/parent::div/following-sibling::div/input"));
            idInp.sendKeys(Keys.CONTROL + "a", Keys.BACK_SPACE);
            idInp.sendKeys(eId);

            if (!photo.isEmpty()) {
                driver.findElement(By.xpath("//input[@type='file']")).sendKeys(photo);
            }
            
            driver.findElement(By.xpath("//button[@type='submit']")).click();

            try {
                // Kiểm tra xem có lưu thành công và chuyển trang không
                wait.until(ExpectedConditions.urlContains("viewPersonalDetails"));
                actRes = "PASS";
                actMsg = "Successfully Saved";
            } catch (Exception e) {
                actRes = "FAIL";
                try {
                    actMsg = driver.findElement(By.xpath("//span[contains(@class,'error')]")).getText();
                } catch (Exception ex) { 
                    actMsg = "Validation failed hoặc lỗi hệ thống"; 
                }
            }
} catch (Exception e) {
            actMsg = "Lỗi: " + e.getMessage();
        }

        // Ghi kết quả vào Sheet
        Row row = sh.getRow(rowIndex);
        row.createCell(colActualRes).setCellValue(actRes);
        row.createCell(colActualMsg).setCellValue(actMsg);
        
        boolean isMatch = actRes.equalsIgnoreCase(expRes);
        if (isMatch) matchedCount++;
        row.createCell(colStatus).setCellValue(isMatch ? "PASS" : "FAIL");

        Assert.assertTrue(isMatch, "TC " + tcID + " không khớp kết quả.");
    }

    @AfterClass
    public void tearDown() throws Exception {
        // GHI BÁO CÁO TỔNG HỢP VÀO CUỐI FILE EXCEL
        double accuracy = (totalTC > 0) ? ((double) matchedCount / totalTC) * 100 : 0;
        accuracy = Math.round(accuracy * 100.0) / 100.0;

        int lastRow = sh.getLastRowNum();
        int summaryRowStart = lastRow + 2;

        String[] labels = {"Tổng số testcase", "Số case khớp (Matched)", "Số case lệch", "Độ chính xác (%)"};
        Object[] values = {totalTC, matchedCount, (totalTC - matchedCount), accuracy + "%"};

        for (int i = 0; i < labels.length; i++) {
            Row r = sh.createRow(summaryRowStart + i);
            r.createCell(0).setCellValue(labels[i]);
            r.createCell(1).setCellValue(values[i].toString());
        }

        // LƯU FILE
        JFileChooser save = new JFileChooser();
        save.setDialogTitle("Lưu file kết quả Excel");
        if (save.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            String path = save.getSelectedFile().getAbsolutePath();
            if (!path.endsWith(".xlsx")) path += ".xlsx";
            FileOutputStream fos = new FileOutputStream(path);
            wb.write(fos);
            fos.close();
            JOptionPane.showMessageDialog(null, "Kiểm thử hoàn tất! Độ chính xác: " + accuracy + "%");
        }

        wb.close();
        if (driver != null) driver.quit();
    }
}