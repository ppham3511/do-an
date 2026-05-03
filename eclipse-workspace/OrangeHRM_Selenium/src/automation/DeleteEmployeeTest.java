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
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class DeleteEmployeeTest implements ITest {
    WebDriver driver;
    WebDriverWait wait;
    JavascriptExecutor js;
    Workbook wb;
    Sheet sh;

    // Cấu hình cột Excel
    int colExpected = 4;   // Cột E
    int colActualRes = 6;  // Cột G
    int colActualMsg = 7;  // Cột H
    int colStatus    = 8;  // Cột I

    int totalTC = 0;
    int matchedCount = 0;

    // Biến ThreadLocal để đặt lại tên hiển thị cho TestNG
    private ThreadLocal<String> testName = new ThreadLocal<>();

    @Override
    public String getTestName() {
        return testName.get();
    }

    @BeforeMethod
    public void setTestName(Method method) {
        testName.set(method.getName()); // Ép giao diện chỉ hiện "testDeleteEmployee"
    }

    @BeforeClass
    public void setup() throws Exception {
        // 1. CHỌN FILE
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Chọn File Excel Input cho Delete Test");
        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
            throw new RuntimeException("Chưa chọn file đầu vào!");
        }

        File inputFile = chooser.getSelectedFile();
        FileInputStream fis = new FileInputStream(inputFile);
        wb = new XSSFWorkbook(fis);
        sh = wb.getSheetAt(0);

        // 2. TẠO TIÊU ĐỀ BÁO CÁO
        Row headerRow = sh.getRow(0);
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        String[] titles = {"Actual Result", "Actual Message", "Status (Match?)"};
        for (int i = 0; i < titles.length; i++) {
            Cell cell = headerRow.createCell(colActualRes + i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(headerStyle);
        }

        // 3. SETUP BROWSER & LOGIN
        ChromeOptions op = new ChromeOptions();
        op.addArguments("--remote-allow-origins=*");
        driver = new ChromeDriver(op);
        driver.manage().window().maximize();
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        js = (JavascriptExecutor) driver;

        loginOrangeHRM();
    }

    @DataProvider(name = "deleteData")
    public Iterator<Object[]> getDeleteData() {
        List<Object[]> data = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        for (int i = 1; i <= sh.getLastRowNum(); i++) {
            Row row = sh.getRow(i);
            if (row == null || df.formatCellValue(row.getCell(0)).isEmpty()) continue;

            data.add(new Object[]{
                df.formatCellValue(row.getCell(1)).trim(), // rawIds
                df.formatCellValue(row.getCell(3)).trim(), // popupAction
                df.formatCellValue(row.getCell(colExpected)).trim().toUpperCase(), // expRes
                i // rowIndex
            });
        }
        return data.iterator();
    }

    @Test(dataProvider = "deleteData")
    public void testDeleteEmployee(String rawIds, String popupAction, String expRes, int rowIndex) {
        totalTC++;
        String actRes = "FAIL";
        String actMsg = "";
        List<String> listTargetIds = Arrays.asList(rawIds.split("\\s*,\\s*"));

        try {
            driver.get("http://localhost/orangehrm/web/index.php/pim/viewEmployeeList");
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("oxd-table-body")));

            js.executeScript("window.scrollTo(0, 500);");
            Thread.sleep(1500);

            int countFound = 0;
            List<WebElement> tableRows = driver.findElements(By.className("oxd-table-card"));

            for (WebElement tableRow : tableRows) {
                String currentId = tableRow.findElement(By.xpath(".//div[@role='cell'][2]")).getText().trim();
                if (listTargetIds.contains(currentId)) {
                    js.executeScript("arguments[0].scrollIntoView({block: 'center'});", tableRow);
                    tableRow.findElement(By.className("oxd-checkbox-input-icon")).click();
                    countFound++;
                }
            }

            if (countFound > 0) {
                WebElement btnDelete = driver.findElement(By.xpath("//button[contains(.,'Delete Selected')]"));
                js.executeScript("arguments[0].click();", btnDelete);

                if (popupAction.equalsIgnoreCase("Yes, Delete")) {
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(.,'Yes, Delete')]"))).click();
                    WebElement toast = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//p[contains(@class,'oxd-text--toast-message')]")));
                    actMsg = "Đã xóa " + countFound + " NV. " + toast.getText();
                    actRes = "PASS";
                } else {
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(.,'No, Cancel')]"))).click();
                    actMsg = "Đã hủy lệnh xóa cho " + countFound + " nhân viên";
                    actRes = "PASS";
                }
            } else {
                actMsg = "Không tìm thấy ID: " + rawIds;
                actRes = "FAIL";
            }
        } catch (Exception e) {
            actRes = "ERROR";
            actMsg = "Lỗi: " + e.getMessage();
        }

        // Ghi kết quả vào Excel
        Row row = sh.getRow(rowIndex);
        row.createCell(colActualRes).setCellValue(actRes);
        row.createCell(colActualMsg).setCellValue(actMsg);
        
        boolean isMatch = actRes.equalsIgnoreCase(expRes);
        if (isMatch) matchedCount++;
        row.createCell(colStatus).setCellValue(isMatch ? "PASS" : "FAIL");

        Assert.assertTrue(isMatch);
    }

    @AfterClass
    public void tearDown() throws Exception {
        // PHẦN THỐNG KÊ TỔNG HỢP
        exportSummaryToExcel();

        // LƯU FILE
        JFileChooser save = new JFileChooser();
        save.setDialogTitle("Lưu Báo Cáo Kết Quả Delete");
        if (save.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            String path = save.getSelectedFile().getAbsolutePath();
            if (!path.endsWith(".xlsx")) path += ".xlsx";
            FileOutputStream fos = new FileOutputStream(path);
            wb.write(fos);
            fos.close();
            JOptionPane.showMessageDialog(null, "Hoàn tất! Báo cáo đã được lưu.");
        }

        wb.close();
        if (driver != null) driver.quit();
    }

    private void loginOrangeHRM() {
        driver.get("http://localhost/orangehrm/web/index.php/auth/login");
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("username"))).sendKeys("admin");
        driver.findElement(By.name("password")).sendKeys("Admin@1234");
        driver.findElement(By.xpath("//button[@type='submit']")).click();
        wait.until(ExpectedConditions.urlContains("dashboard"));
    }

    private void exportSummaryToExcel() {
        int summaryRowIdx = sh.getLastRowNum() + 2;
        double accuracy = (totalTC > 0) ? ((double) matchedCount / totalTC) * 100 : 0;
        accuracy = Math.round(accuracy * 100.0) / 100.0;

        String[] labels = {"Tổng số testcase", "Số case khớp (Matched)", "Số case lệch", "Độ chính xác (%)"};
        Object[] values = {totalTC, matchedCount, (totalTC - matchedCount), accuracy + "%"};

        CellStyle style = wb.createCellStyle();
        Font f = wb.createFont(); f.setBold(true); style.setFont(f);

        for (int j = 0; j < labels.length; j++) {
            Row r = sh.createRow(summaryRowIdx + j);
            Cell c = r.createCell(0);
            c.setCellValue(labels[j]);
            c.setCellStyle(style);
            r.createCell(1).setCellValue(values[j].toString());
        }
    }
}