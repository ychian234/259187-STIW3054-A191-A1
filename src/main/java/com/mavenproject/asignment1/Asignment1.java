package com.mavenproject.asignment1;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
public class Asignment1 {
    private static final List<ExcelData> dataRecord = new ArrayList();
    private static final String URL = "https://github.com/STIW3054-A191/Assignments/wiki/List_of_Student";
    private static final String URL2 = "https://github.com/STIW3054-A191/Main-Issues/issues/1";
    public static void getGitData() {
        try {
            System.out.println("Accessing " + URL + "...");
            Document doc1 = Jsoup.connect(URL).get();
            Element table = doc1.select("table").get(0);
            Elements rows = table.select("tr");
            for (Element row : rows) {
                Elements th1 = row.select("th:nth-child(1)");
                Elements th2 = row.select("th:nth-child(2n)");
                Elements th3 = row.select("th:nth-child(3n)");
                Elements data1 = row.select("td:nth-child(1)");
                Elements data2 = row.select("td:nth-child(2n)");
                Elements data3 = row.select("td:nth-child(3n)");
                String header1 = th1.text();
                String header2 = th2.text();
                String header3 = th3.text();
                String column1 = data1.text();
                String column2 = data2.text();
                String column3 = data3.text();
                dataRecord.add(new ExcelData(header1 + column1, header2 + column2, header3 + column3));
            }
            System.out.println("Table data has been collected successfully.");
            System.out.println();
        } catch (IOException e) {
            System.out.println("ERROR : Failed to access " + URL);
        }
    }
    public static void writeToExcel() {
        if (dataRecord.isEmpty()) {
            System.out.println("ERROR : No data to write, build terminated.");
            System.exit(0);
        }
        String excelFile = "STIW3054_A191.xlsx";
        System.out.println("Writing the " + excelFile + "...");
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Student_List_Total");
        XSSFSheet sheet2 = workbook.createSheet("Student_Not_Comment");
        try {
            for (int i = 0; i < dataRecord.size(); i++) {
                XSSFRow row = sheet.createRow(i);
                XSSFCell cell1 = row.createCell(0);
                cell1.setCellValue(dataRecord.get(i).getData0());
                XSSFCell cell2 = row.createCell(1);
                cell2.setCellValue(dataRecord.get(i).getData1());
                XSSFCell cell3 = row.createCell(2);
                cell3.setCellValue(dataRecord.get(i).getData2());
                if (i > 0) {
                    Document doc2 = Jsoup.connect(URL2).get();
                    Elements body1 = doc2.select("td");
                    Elements comments = body1.select("p:contains(" + dataRecord.get(i).getData1() + ")");
                    Elements links = comments.select("a");
                    String glinks = links.text();
                    XSSFCell cell4 = row.createCell(3);
                    cell4.setCellValue(glinks);
                } else {
                    XSSFCell cell4 = row.createCell(3);
                    cell4.setCellValue("Github Link");
                }
            }
            Document doc2 = Jsoup.connect(URL2).get();
            Elements body1 = doc2.select("td");
            Elements comments = body1.select("p");
            String temp = comments.text();
            int k = 0;
            do {
                if (k == 0) {
                    XSSFRow row = sheet2.createRow(k);
                    sheet2.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
                    XSSFCell cell1 = row.createCell(0);
                    cell1.setCellValue("Students that do not comments in github.");
                    k++;
                } else if (k == 1) {
                    XSSFRow row = sheet2.createRow(k);
                    XSSFCell cell2 = row.createCell(0);
                    cell2.setCellValue("No.");
                    XSSFCell cell3 = row.createCell(1);
                    cell3.setCellValue("Matric");
                    XSSFCell cell4 = row.createCell(2);
                    cell4.setCellValue("Name");
                    k++;
                } else {
                    int l = 2;
                    for (int i = 0; i < dataRecord.size(); i++) {
                        if (!temp.contains(dataRecord.get(i).getData1())) {
                            XSSFRow row = sheet2.createRow(l);
                            XSSFCell cell5 = row.createCell(0);
                            cell5.setCellValue(l - 1);
                            XSSFCell cell6 = row.createCell(1);
                            cell6.setCellValue(dataRecord.get(i).getData1());
                            XSSFCell cell7 = row.createCell(2);
                            cell7.setCellValue(dataRecord.get(i).getData2());
                            l++;
                        }
                    }
                    k++;
                }
            } while (k < 3);
            FileOutputStream outputFile = new FileOutputStream(excelFile);
            workbook.write(outputFile);
            outputFile.flush();
            outputFile.close();
            System.out.println(excelFile + " Is written successfully.");
        } catch (IOException e) {
            System.out.println("ERROR : Failed to write the file!");
        }
    }

    public static void main(String[] args) {
        getGitData();
        writeToExcel();
    }
}