package com.example.demo.service;

import com.example.demo.model.FileItem;
import com.example.demo.model.UserEntity;
import com.example.demo.repository.CustomUserRepository;
import com.example.demo.repository.UserRepository;
import com.example.demo.utils.DbReaderTask;
import com.example.demo.utils.ExcelUtils;
import com.example.demo.utils.ExcelWriterTask;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.LinkedBlockingQueue;

@Service
@Slf4j
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class UserService {

    static int QUERY_BATCH_SIZE = 1000;
    UserRepository userRepository;
    CustomUserRepository customUserRepository;
    ExecutorService executorService;

    public FileItem exportFromDbToFile() {
        String filePath = String.format("User_Report_%s.xlsx",
                new SimpleDateFormat("ddMMyyyy").format(new Date()));
        final String title = "USER_REPORT";
        List<String> headers = ExcelUtils.extractHeadersFromObjects(UserEntity.class);
        List<String> fileDescriptions = new ArrayList<>();
        fileDescriptions.add("From: dd/MM/yyyy" + " - To: dd/MM/yyyy");

        ExcelUtils.generateFileHeader(filePath, title, fileDescriptions, headers);

        BlockingQueue<List<UserEntity>> queue = new LinkedBlockingQueue<>(10);

        CompletableFuture<Void> readFuture = CompletableFuture
                .runAsync(new DbReaderTask<>(userRepository, queue, QUERY_BATCH_SIZE), executorService);

        CompletableFuture<Void> writeFuture = CompletableFuture
                .runAsync(new ExcelWriterTask<>(queue, filePath, ExcelUtils::mapObjectToRow), executorService);

        CompletableFuture<Void> combinedFuture = CompletableFuture.allOf(readFuture, writeFuture);

        combinedFuture.thenRun(() -> log.info("Completed write to file: {}", filePath)).join();

        return ExcelUtils.loadFileAsResource(filePath);
    }

    public Map<String, Object> getDataFromExcel(InputStream inputStream, int headerHeight) throws IOException {
        List<UserEntity> users = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        UserEntity user;

        int rowIndex = 0;
        Map<Integer, String> errorMessageMapByRows = new HashMap<>();
        for (Row row : sheet) {
            if (rowIndex < headerHeight) {
                rowIndex++;
                continue;
            }

            user = new UserEntity();
            try {
                Iterator<Cell> cellIterator = row.iterator();
                int cellIndex = 0;
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cellIndex) {
                        case 1 -> user.setId((long) cell.getNumericCellValue());
                        case 2 -> user.setUsername(cell.getStringCellValue());
                        case 3 -> user.setEmail(cell.getStringCellValue());
                        case 4 -> user.setPhone(cell.getStringCellValue());
                        case 5 -> user.setAge((int) cell.getNumericCellValue());
                        default -> {}
                    }
                    cellIndex++;
                }
                users.add(user);
            } catch (Exception e) {
                String errorMessage = String.format(e.getMessage());
                errorMessageMapByRows.put(row.getRowNum() + 1, errorMessage);
            }
            rowIndex++;
        }
        Map<String, Object> response = new HashMap<>();
        response.put("users", users);
        response.put("errorMessageMapByRows", errorMessageMapByRows);
        return response;
    }

    public FileItem saveWithErrorReport(MultipartFile file) throws IOException {
        if (!ExcelUtils.hasExcelFormat(file)) {
            throw new RuntimeException("Invalid file format");
        }

        Map<String, Object> data = getDataFromExcel(file.getInputStream(), 3);
        List<UserEntity> users = (List<UserEntity>) data.get("users");
        Map<Integer, String> errorMessageMapByRows = (Map<Integer, String>) data.get("errorMessageMapByRows");
        customUserRepository.saveAll(users);

        if (!errorMessageMapByRows.isEmpty()) {
            String errorFilePath = String.format("Error_Report_%s.xlsx",
                    new SimpleDateFormat("ddMMyyyy").format(new Date()));
            generateErrorFile(errorFilePath, errorMessageMapByRows);
            return ExcelUtils.loadFileAsResource(errorFilePath);
        }

        return new FileItem();
    }

    private void generateErrorFile(String filePath, Map<Integer, String> errorMessageMapByRows) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Error Report");
            int rowNum = 0;

            Row headerRow = sheet.createRow(rowNum++);
            headerRow.createCell(0).setCellValue("Row Num");
            headerRow.createCell(1).setCellValue("Error Message");

            for (Integer key : errorMessageMapByRows.keySet()) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(key);
                row.createCell(1).setCellValue(errorMessageMapByRows.get(key));
            }

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
            log.info("Error report file written successfully at: {}", filePath);
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
    }

}
