package com.example;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
