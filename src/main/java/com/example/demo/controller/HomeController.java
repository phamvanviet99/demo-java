package com.example.demo.controller;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;


import org.springframework.core.io.Resource;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

import java.io.IOException;
import java.io.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import net.lingala.zip4j.ZipFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;


@Controller
public class HomeController {

    @GetMapping("/")
    public String index(Model model) {
        model.addAttribute("message", "Xin chào");
        return "index"; // trả về file templates/index.html
    }



    @PostMapping("/process")
    public void processFileStream(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        Path outputDir = null;
        try {
            outputDir = Files.createTempDirectory("task2-"); // short-lived temp
            generateReports(file, outputDir);
            final Path finalOutputDir = outputDir;


            response.setStatus(HttpServletResponse.SC_OK);
            response.setContentType("application/zip");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"result.zip\"");

            try (ServletOutputStream sos = response.getOutputStream();
                 ZipOutputStream zos = new ZipOutputStream(new BufferedOutputStream(sos))) {

                Files.walk(finalOutputDir)
                        .filter(p -> !Files.isDirectory(p))
                        .forEach(p -> {
                            String entryName = finalOutputDir.relativize(p).toString().replace("\\", "/");
                            ZipEntry entry = new ZipEntry(entryName);
                            try (InputStream is = Files.newInputStream(p)) {
                                zos.putNextEntry(entry);
                                byte[] buffer = new byte[8192];
                                int len;
                                while ((len = is.read(buffer)) != -1) {
                                    zos.write(buffer, 0, len);
                                }
                                zos.closeEntry();
                            } catch (IOException e) {
                                throw new UncheckedIOException(e);
                            }
                        });
                zos.finish();
            }

        } catch (Exception e) {
            e.printStackTrace();
            try {
                response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Error when creating zip: " + e.getMessage());
            } catch (IOException ignored) {}
        } finally {
            // xóa temp outputDir nếu cần
            if (outputDir != null) {
                try { Files.walk(outputDir).sorted(Comparator.reverseOrder()).map(Path::toFile).forEach(File::delete); }
                catch (IOException ignored) {}
            }
        }
    }

    private void generateReports(MultipartFile file, Path outputDir) throws Exception {
        // 🔹 Tạo file tạm để đọc dữ liệu upload
        Path tempFile = Files.createTempFile("summary-upload-", ".xlsx");
        file.transferTo(tempFile.toFile());

        // 🔹 Kiểm tra template có tồn tại local hay trong resources
        boolean useClasspathTemplate = false;
        File templateFile = new File("./templates/Template2.xlsx");
        if (!templateFile.exists()) {
            useClasspathTemplate = true;
            System.out.println("📦 Template không có ở local, dùng file trong resources.");
        } else {
            System.out.println("📁 Template local: " + templateFile.getAbsolutePath());
        }

        try (FileInputStream fis = new FileInputStream(tempFile.toFile());
             Workbook summaryWb = new XSSFWorkbook(fis)) {

            Sheet summarySheet = summaryWb.getSheet("Báo cáo - Gửi đại lý");
            if (summarySheet == null) {
                throw new IllegalArgumentException("Không tìm thấy sheet 'Báo cáo - Gửi đại lý' trong file upload");
            }

            int startRow = 6;
            int endRow = 55;
            int blockWidth = 12;
            int startCol = 2;

            Row headerRow = summarySheet.getRow(4);
            int lastCol = headerRow.getLastCellNum();

            // 🔹 Ngày cuối tháng trước
            LocalDate today = LocalDate.now();
            YearMonth lastMonth = YearMonth.from(today.minusMonths(1));
            LocalDate lastDateOfPrevMonth = lastMonth.atEndOfMonth();

            int currentMonth = today.getMonthValue();
            int lastMonth1 = currentMonth - 1;
            if (lastMonth1 < 1) lastMonth1 = 12;
            int dem = 0;

            // 🔹 Lặp qua từng cửa hàng
            for (int col = startCol; col < lastCol; col += blockWidth) {
                Cell cell = headerRow.getCell(col);
                if (cell == null || cell.toString().trim().isEmpty()) continue;

                String storeName = cell.toString().trim();
                System.out.println("🏪 Exporting store: " + storeName);

                // ⚙️ 👉 Mỗi lần lặp: mở lại InputStream mới cho template
                try (InputStream templateStream = useClasspathTemplate
                        ? new org.springframework.core.io.ClassPathResource("templates/Template2.xlsx").getInputStream()
                        : new FileInputStream(templateFile);
                     Workbook templateWb = new XSSFWorkbook(templateStream)) {

                    Sheet templateSheet = templateWb.getSheetAt(0);

                    // 🔸 Ghi D3 (ngày cuối tháng trước)
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
                    String formattedDate = lastDateOfPrevMonth.format(formatter);
                    Row rowB3 = templateSheet.getRow(2);
                    if (rowB3 == null) rowB3 = templateSheet.createRow(2);
                    Cell cellB3 = rowB3.getCell(3);
                    if (cellB3 == null) cellB3 = rowB3.createCell(3);
                    cellB3.setCellValue(formattedDate);

                    // 🔸 Copy dữ liệu
                    for (int r = startRow; r <= endRow; r++) {
                        Row srcRow = summarySheet.getRow(r);
                        Row destRow = templateSheet.getRow(r - startRow + 5);
                        if (destRow == null) destRow = templateSheet.createRow(r - startRow + 5);

                        double sum = 0;
                        for (int c = 0; c < blockWidth; c++) {
                            dem++;
                            Cell srcCell = srcRow.getCell(col + c);
                            Cell destCell = destRow.getCell(c + 4);
                            if (destCell == null) destCell = destRow.createCell(c + 4);

                            if (srcCell != null) {
                                copyCellValue(srcCell, destCell);
                                if (srcCell.getCellType() == CellType.NUMERIC && dem <= lastMonth1) {
                                    sum += srcCell.getNumericCellValue();
                                }
                            }
                        }
                        dem = 0;

                        if (r < 49) {
                            Cell sumCell = destRow.getCell(16);
                            if (sumCell == null) sumCell = destRow.createCell(16);
                            sumCell.setCellValue(sum);
                        }
                    }

                    // 🔸 Ghi mã đại lý (D2)
                    String branchCode = "";
                    Sheet danhMucSheet = summaryWb.getSheet("Danh mục");
                    if (danhMucSheet != null) {
                        for (int i = 1; i <= danhMucSheet.getLastRowNum(); i++) {
                            Row row = danhMucSheet.getRow(i);
                            if (row == null) continue;

                            Cell nameCell = row.getCell(3);
                            Cell codeCell = row.getCell(2);

                            if (nameCell != null && codeCell != null) {
                                String name = nameCell.toString().trim();
                                if (name.equalsIgnoreCase(storeName)) {
                                    branchCode = codeCell.toString().trim();
                                    break;
                                }
                            }
                        }
                    }

                    if (branchCode.isEmpty()) {
                        branchCode = storeName.replaceAll("[\\\\/:*?\"<>|]", "_");
                    }

                    Row rowB2 = templateSheet.getRow(1);
                    if (rowB2 == null) rowB2 = templateSheet.createRow(1);
                    Cell cellB2 = rowB2.getCell(3);
                    if (cellB2 == null) cellB2 = rowB2.createCell(3);
                    cellB2.setCellValue(branchCode);

                    // 🔸 Xuất file
                    int year = today.getYear();
                    String fileName = String.format("BCKQKD_%02dT%d_%s.xlsx", lastMonth1, year, branchCode);
                    Path outputFile = outputDir.resolve(fileName);

                    try (FileOutputStream fos = new FileOutputStream(outputFile.toFile())) {
                        templateWb.write(fos);
                    }

                    System.out.println("✅ Exported: " + outputFile);
                }
            }
        }

        Files.deleteIfExists(tempFile);
    }




    private static void copyCellValue(Cell src, Cell dest) {
        switch (src.getCellType()) {
            case STRING:
                dest.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                dest.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dest.setCellValue(src.getBooleanCellValue());
                break;
            case FORMULA:
                dest.setCellFormula(src.getCellFormula());
                break;
            case BLANK:
                dest.setBlank();
                break;
            default:
                dest.setCellValue(src.toString());
                break;
        }
    }

    private static void createFolderIfNotExists(String folderPath) throws IOException {
        Path path = Paths.get(folderPath);
        if (!Files.exists(path)) {
            Files.createDirectories(path);
            System.out.println("📁 Đã tạo thư mục: " + folderPath);
        } else {
            System.out.println("📂 Thư mục đã tồn tại: " + folderPath);
        }
    }

}