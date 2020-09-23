package com.example.demo.File;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.*;
import java.nio.file.Files;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Stream;

@Controller
public class FileUploadController {


    @GetMapping("/")
    public String homepage() throws IOException {
        return "index";
    }

    @PostMapping("/convert")
    public ResponseEntity<Resource> convertCSVFileTOXLS(@RequestParam("file") MultipartFile file, RedirectAttributes attributes) throws Exception {

        HttpHeaders header = new HttpHeaders();

        if (file.isEmpty()) {
            header.add("Location", "/");
            attributes.addFlashAttribute("message", "Выберите файл для загрузки.");
            return new ResponseEntity<Resource>(null, header, HttpStatus.FOUND);
        }

        String fileName = StringUtils.cleanPath(file.getOriginalFilename());
        if (!getFileExtension(fileName).equals("csv")) {
            header.add("Location", "/");
            attributes.addFlashAttribute("message", "Пожалуйста, выберите правильный файл.");
            return new ResponseEntity<Resource>(null, header, HttpStatus.FOUND);
        }

        File file1 = multipartToFile(file, fileName);

        String sheetText = fileName.replaceFirst("[.][^.]+$", "");
        String newXlSFilename = sheetText + ".xls";

        attributes.addFlashAttribute("message", "Файл конвертировался успешно" + fileName + '!');


        ByteArrayResource resource = new ByteArrayResource(convertCsvToXlsx(file1, sheetText));

        header.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + newXlSFilename);
        header.add("Cache-Control", "no-cache, no-store, must-revalidate");
        header.add("Pragma", "no-cache");
        header.add("Expires", "0");

        return ResponseEntity.ok()
                .headers(header)
                .contentLength(convertCsvToXlsx(file1, sheetText).length)
                .contentType(MediaType.parseMediaType("application/octet-stream"))
                .body(resource);

    }


    public static byte[] convertCsvToXlsx(File csvLocation, String sheetText) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetText);
        AtomicReference<Integer> row = new AtomicReference<>(0);
        Files.readAllLines(csvLocation.toPath()).forEach(line -> {
            Row currentRow = sheet.createRow(row.getAndSet(row.get() + 1));
            String[] nextLine = line.split(",");
            Stream.iterate(0, i -> i + 1).limit(nextLine.length).forEach(i -> {
                currentRow.createCell(i).setCellValue(nextLine[i]);
            });
        });
        ByteArrayOutputStream fos = new ByteArrayOutputStream();
        workbook.write(fos);
        byte[] xlsArray = fos.toByteArray();
        fos.flush();
        return xlsArray;
    }

    private static String getFileExtension(String fileName) {
        if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
            return fileName.substring(fileName.lastIndexOf(".") + 1);
        else return "";
    }

    public static File multipartToFile(MultipartFile multipart, String fileName) throws IllegalStateException, IOException {
        File convFile = new File(System.getProperty("java.io.tmpdir") + "/" + fileName);
        multipart.transferTo(convFile);
        return convFile;
    }

}
