package scheduler;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.time.Duration;

import static org.openqa.selenium.support.ui.ExpectedConditions.*;

public class AmazonCitywise_1 {

    private static String lastAppliedPincode = null;
    private static WebDriverWait fastWait;
    private static WebDriverWait normalWait;

    public static void main(String[] args) {

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-images", "--disable-extensions", "--no-sandbox",
                "--disable-dev-shm-usage", "--disable-gpu", "--blink-settings=imagesEnabled=false");
        options.setPageLoadStrategy(PageLoadStrategy.EAGER);
        // options.addArguments("--headless=new");

        WebDriver driver = new ChromeDriver(options);
        driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(40));
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0)); // We use explicit waits only

        normalWait = new WebDriverWait(driver, Duration.ofSeconds(15));
        fastWait   = new WebDriverWait(driver, Duration.ofSeconds(3));

        try {
            String filePath = ".\\input-data\\CityWise300 Input DataUpdatedNew.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook wb = new XSSFWorkbook(file);
            Sheet sheet = wb.getSheet("Amazon1");

            List<String[]> data = new ArrayList<>();
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String url = getCellValue(row.getCell(5));
                if (url.isEmpty() || url.equalsIgnoreCase("NA")) continue;

                data.add(new String[]{
                    getCellValue(row.getCell(0)),
                    getCellValue(row.getCell(1)),
                    getCellValue(row.getCell(2)),
                    getCellValue(row.getCell(3)),
                    getCellValue(row.getCell(4)),
                    url,
                    getCellValue(row.getCell(6)),
                    getCellValue(row.getCell(7)),
                    getCellValue(row.getCell(9)),
                    getCellValue(row.getCell(10))
                });
            }
            wb.close(); file.close();

            Workbook outWb = new XSSFWorkbook();
            Sheet outSheet = outWb.createSheet("Results");
            String[] headers = {"InputPid","InputCity","InputName","InputSize","NewProductCode","URL","Name","MRP","SP",
                    "UOM","Multiplier","Availability","Offer","Commands","Remarks","Correctness","Percentage","Name","Name Check"};
            Row h = outSheet.createRow(0);
            for (int i = 0; i < headers.length; i++) h.createCell(i).setCellValue(headers[i]);

            int rowIdx = 1;

            for (int i = 0; i < data.size(); i++) {
                String[] d = data.get(i);
                String id = d[0], city = d[1], name = d[2], size = d[3], code = d[4], url = d[5],
                       uom = d[6], mult = d[7], pin = d[8], namecheck = d[9];

                String productName = "NA", mrp = "NA", sp = "NA", offer = "NA", avail = "1";

                try {
                    driver.get(url);
                    setPincodeSmart(driver, pin.trim());

                    productName = getTextOrNA(By.id("productTitle"));
                    sp = getTextOrNA(By.xpath("(//span[contains(@class,'a-price-whole') and ancestor::span[contains(@class,'priceToPay')]])[1]"))
                            .replace("₹","").replace(",","");

                    String mrpRaw = getTextOrNA(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[2]/span/span[1]/span[2]/span"));
                    if (!mrpRaw.equals("NA")) {
                        mrp = mrpRaw.replace("₹","").replace(",","");
                    } else if (!sp.equals("NA")) {
                        mrp = sp;
                    }

                    String stock = getTextOrNA(By.id("availability")).toLowerCase();
                    avail = (stock.contains("in stock") || stock.contains("only")) ? "0" : "1";

                    if (!mrp.equals(sp) && !sp.equals("NA")) {
                        offer = getTextOrNA(By.xpath("//span[contains(@class,'savingsPercentage')]"))
                                .replace("-","").replace("(","").replace(")","");
                    }

                    // FIXED: FULLY CORRECT ROW WRITING
                    Row r = outSheet.createRow(rowIdx++);
                    r.createCell(0).setCellValue(id);
                    r.createCell(1).setCellValue(city);
                    r.createCell(2).setCellValue(name);
                    r.createCell(3).setCellValue(size);
                    r.createCell(4).setCellValue(code);
                    r.createCell(5).setCellValue(url);
                    r.createCell(6).setCellValue(productName);
                    r.createCell(7).setCellValue(mrp);
                    r.createCell(8).setCellValue(sp);
                    r.createCell(9).setCellValue(uom);
                    r.createCell(10).setCellValue(mult);
                    r.createCell(11).setCellValue(avail);
                    r.createCell(12).setCellValue(offer);
                    r.createCell(18).setCellValue(namecheck);

                    System.out.println("DONE [" + (i+1) + "/" + data.size() + "] " + productName);

                } catch (Exception e) {
                    writeNA(outSheet, rowIdx++, id, city, name, size, code, url, uom, mult, namecheck);
                    System.out.println("FAILED [" + (i+1) + "] " + url);
                }
            }

            String outPath = ".\\Output\\CityWise_Amazon_1_" + 
                new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xlsx";
            FileOutputStream fos = new FileOutputStream(outPath);
            outWb.write(fos); fos.close(); outWb.close();

            System.out.println("COMPLETED! Saved: " + outPath);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }

    private static String getTextOrNA(By locator) {
        try {
            return fastWait.until(visibilityOfElementLocated(locator)).getText().trim();
        } catch (TimeoutException e) {
            return "NA";
        }
    }

 // GET CELL VALUE SAFELY
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        try {
            if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue().trim();
            if (cell.getCellType() == CellType.NUMERIC) return String.valueOf((long)cell.getNumericCellValue());
        } catch (Exception e) { }
        return "";
    }

    private static void writeNA(Sheet s, int r, String... v) {
        Row row = s.createRow(r);
        for (int i = 0; i < v.length; i++) row.createCell(i).setCellValue(i < v.length ? v[i] : "");
        row.createCell(6).setCellValue("NA"); row.createCell(7).setCellValue("NA"); row.createCell(8).setCellValue("NA");
    }

    private static void setPincodeSmart(WebDriver driver, String targetPin) {
        if (targetPin == null || targetPin.isEmpty() || targetPin.equals(lastAppliedPincode)) return;

        try {
            normalWait.until(elementToBeClickable(By.id("glow-ingress-block"))).click();
            WebElement input = normalWait.until(elementToBeClickable(By.id("GLUXZipUpdateInput")));
            input.clear(); input.sendKeys(targetPin);
            normalWait.until(elementToBeClickable(By.xpath("//span[contains(text(),'Apply')]/parent::*"))).click();
            normalWait.until(d -> driver.findElement(By.id("glow-ingress-block")).getText().contains(targetPin));
            lastAppliedPincode = targetPin;
            System.out.println("Pincode set: " + targetPin);
        } catch (Exception e) {
            lastAppliedPincode = targetPin;
        }
    }

}
