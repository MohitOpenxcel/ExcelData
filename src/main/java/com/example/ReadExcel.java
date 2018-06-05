package com.example;

import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ReadExcel {

    @Value("${application.message:Hello World}")
    private String message = "Hello World";

    @Autowired
    private UserService userService;

    @RequestMapping(method = RequestMethod.POST, value = "/excel/template")
    public void generateXSLTemplate() throws IOException {

        final String FILE_NAME = "C:\\Users\\Mohit\\Desktop\\data2.xlsx";
        String[] columns = {"name", "username", "password", "email"};

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sheet2");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("NAME");
        header.createCell(1).setCellValue("AGE");
        header.createCell(2).setCellValue("GENDER");
        header.createCell(3).setCellValue("PASSWORD");

        CellRangeAddressList addressList = new CellRangeAddressList(1, 50, 2, 2);
        DVConstraint dvConstraint = DVConstraint
                .createExplicitListConstraint(new String[] { "male", "female"});
        DataValidation dataValidation = new HSSFDataValidation(addressList,
                dvConstraint);

        dataValidation.setSuppressDropDownArrow(false);
        sheet.addValidationData(dataValidation);

        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }

    @RequestMapping(method = RequestMethod.POST, value = "/excel/add")
    public void addXLSDataTODb(@RequestParam("file") MultipartFile file) throws IOException {
        if (!file.isEmpty()) {
            InputStream excelFile = file.getInputStream();
            XSSFWorkbook workbook = new XSSFWorkbook(excelFile);

            for (int k = 0; k < workbook.getNumberOfSheets(); k++) {
                Sheet sheet = workbook.getSheetAt(k);
                Iterator<Row> rows = sheet.iterator();
                List list;
                while (rows.hasNext()) {
                    Row currentRow = rows.next();
                    Iterator<Cell> cellsInRow = currentRow.iterator();
                    list = new ArrayList();
                    User user = new User();

                    while (cellsInRow.hasNext()) {

                        Cell currentCell = cellsInRow.next();
                        list.add(currentCell);

                    }
                    user.setName(String.valueOf(list.get(0)));
                    user.setUsername(String.valueOf(list.get(1)));
                    user.setPassword(String.valueOf(list.get(2)));
                    user.setEmail(String.valueOf(list.get(3)));
                    userService.addUser(user);
                }

                workbook.close();

                try {
                    excelFile.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}
