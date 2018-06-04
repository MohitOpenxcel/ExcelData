package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ReadExcel {

    @Value("${application.message:Hello World}")
    private String message = "Hello World";

    @Autowired
    private UserService userService;

     // @RequestMapping(method = RequestMethod.POST, value = "/excel/add")
     public void addXLSDataTODb() throws IOException {

        FileInputStream excelFile = new FileInputStream(new File("C:\\Users\\Mohit\\Desktop\\data.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(excelFile);

        for(int k=0; k<workbook.getNumberOfSheets(); k++){
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
        }

         excelFile.close();
    }

}