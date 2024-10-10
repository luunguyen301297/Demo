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
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
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

    public List<UserEntity> getDataFromExcel(InputStream inputStream, int headerHeight) throws IOException {
        List<UserEntity> users = new ArrayList<>();
        XSSFWorkbook workbook;
        workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        UserEntity user;
        int rowIndex = 0;
        for (Row row : sheet) {
            if (rowIndex < headerHeight) {
                rowIndex++;
                continue;
            }
            Iterator<Cell> cellIterator = row.iterator();
            int cellIndex = 0;
            user = new UserEntity();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cellIndex) {
                    case 1 -> user.setId((long) cell.getNumericCellValue());
                    case 2 -> user.setUsername(cell.getStringCellValue());
                    case 3 -> user.setEmail(cell.getStringCellValue());
                    case 4 -> user.setPhone(cell.getStringCellValue());
                    case 5 -> user.setAge((int) cell.getNumericCellValue());
                    default -> {
                    }
                }
                cellIndex++;
            }
            users.add(user);
        }

        return users;
    }

    public void save(MultipartFile file) throws IOException {
        if (!ExcelUtils.hasExcelFormat(file)) throw new RuntimeException("err format");
        List<UserEntity> users = this.getDataFromExcel(file.getInputStream(), 3);
        customUserRepository.saveAll(users);
    }

    public FileItem saveWithErrorReport(MultipartFile file) throws IOException {
        if (!ExcelUtils.hasExcelFormat(file)) {
            throw new RuntimeException("Invalid file format");
        }

        List<UserEntity> users = this.getDataFromExcel(file.getInputStream(), 3);
        List<UserEntity> failedUsers = new ArrayList<>();
        List<String> errorMessages = new ArrayList<>();

        for (UserEntity user : users) {
            try {
                userRepository.save(user);
            } catch (Exception e) {
                failedUsers.add(user);
                errorMessages.add(e.getMessage());
            }
        }

        if (!failedUsers.isEmpty()) {
            String errorFilePath = String.format("Error_Report_%s.xlsx",
                    new SimpleDateFormat("ddMMyyyy").format(new Date()));
            generateErrorFile(errorFilePath, failedUsers, errorMessages);
            return ExcelUtils.loadFileAsResource(errorFilePath);
        }

        return null;
    }

    private void generateErrorFile(String filePath, List<UserEntity> failedUsers, List<String> errorMessages) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Error Report");
            int rowNum = 0;

            Row headerRow = sheet.createRow(rowNum++);
            headerRow.createCell(0).setCellValue("ID");
            headerRow.createCell(1).setCellValue("Username");
            headerRow.createCell(2).setCellValue("Email");
            headerRow.createCell(3).setCellValue("Phone");
            headerRow.createCell(4).setCellValue("Age");
            headerRow.createCell(5).setCellValue("Error Message");

            for (int i = 0; i < failedUsers.size(); i++) {
                UserEntity user = failedUsers.get(i);
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getUsername());
                row.createCell(2).setCellValue(user.getEmail());
                row.createCell(3).setCellValue(user.getPhone());
                row.createCell(4).setCellValue(user.getAge());
                row.createCell(5).setCellValue(errorMessages.get(i));
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
