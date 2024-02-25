package com.converter;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.opencsv.CSVWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;
import java.util.Map;

class Converter {

    public static void xlsxToJson(String basePath) {
        try {
            // Define the relative path to the input Excel file
            String excelFileName = findFileName(basePath, ".xlsx");
            if (excelFileName == null) {
                System.err.println("No .xlsx file found in the directory: " + basePath);
                return;
            }

            // Construct the absolute path using Paths.get()
            String excelFilePath = basePath + File.separator + excelFileName;
            System.out.println(excelFilePath);

            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);

            ObjectMapper objectMapper = new ObjectMapper();
            ObjectNode root = objectMapper.createObjectNode();

            // Iterate through each sheet in the workbook
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                ArrayNode jsonArray = objectMapper.createArrayNode();

                // Fetch column names from the first row (header row)
                Row headerRow = sheet.getRow(0);
                String[] columnNames = new String[headerRow.getLastCellNum()];
                for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                    columnNames[j] = headerRow.getCell(j).getStringCellValue().trim();
                }

                // Iterate through each row (skipping header row)
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    ObjectNode rowObject = objectMapper.createObjectNode();

                    // Iterate through each cell in the row
                    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        String columnName = columnNames[cellIndex];
                        rowObject.put(columnName, cellToString(cell));
                    }

                    jsonArray.add(rowObject);
                }

                // Add JSON array for current sheet
                root.set(sheet.getSheetName(), jsonArray);
            }

            // Construct the output JSON file name based on the input Excel file name
            String jsonFileName = excelFileName.replace(".xlsx", ".json");
            String jsonFilePath = basePath + File.separator + jsonFileName;

            // Write JSON object to file with UTF-8 BOM
            try (OutputStreamWriter writer = new OutputStreamWriter(new FileOutputStream(jsonFilePath), StandardCharsets.UTF_8)) {
                writer.write('\ufeff'); // Write UTF-8 BOM
                objectMapper.writerWithDefaultPrettyPrinter().writeValue(writer, root);
            }

            workbook.close();
            inputStream.close();

            System.out.println("Conversion completed successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    public static void jsonToCsv(String basePath) {
        try {
            String jsonFileName = findFileName(basePath, ".json");
            if (jsonFileName == null) {
                System.err.println("No .json file found in the directory: " + basePath);
                return;
            }

            String jsonFilePath = basePath + File.separator + jsonFileName;


            // Check if JSON file exists
            File jsonFile = new File(jsonFilePath);
            if (jsonFile.exists()) {
                // Read JSON file
                ObjectMapper objectMapper = new ObjectMapper();
                JsonNode jsonNode = objectMapper.readTree(jsonFile);

                // Extract directory path from JSON file path
                String directoryPath = jsonFile.getParent();

                // Iterate over JSON object fields (arrays)
                Iterator<String> fieldNames = jsonNode.fieldNames();
                while (fieldNames.hasNext()) {
                    String arrayName = fieldNames.next();
                    JsonNode arrayData = jsonNode.get(arrayName);

                    // Output CSV file path with array name
                    String csvFilePath = directoryPath + File.separator + arrayName + ".csv";

                    // Create CSV writer for current array
                    CSVWriter csvWriter = new CSVWriter(new FileWriter(csvFilePath));

                    // Write CSV headers
                    JsonNode firstRow = arrayData.get(0);
                    Iterator<String> fieldNamesIterator = firstRow.fieldNames();
                    String[] headers = new String[firstRow.size()];
                    int index = 0;
                    while (fieldNamesIterator.hasNext()) {
                        headers[index++] = fieldNamesIterator.next();
                    }
                    csvWriter.writeNext(headers);

                    // Write data rows to CSV
                    for (JsonNode row : arrayData) {
                        Iterator<JsonNode> elements = row.elements();
                        String[] rowData = new String[row.size()];
                        int i = 0;
                        while (elements.hasNext()) {
                            rowData[i++] = elements.next().asText();
                        }
                        csvWriter.writeNext(rowData);
                    }

                    // Close CSV writer for current array
                    csvWriter.close();

                    System.out.println("CSV file for array '" + arrayName + "' has been created successfully!");
                }
            } else {
                System.out.println("No JSON file found at: " + jsonFilePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    public static void jsonToXlsx(String basePath) {
        try {
            String jsonFileName = findFileName(basePath, ".json");
            if (jsonFileName == null) {
                System.err.println("No .json file found in the directory: " + basePath);
                return;
            }

            String jsonFilePath = basePath + File.separator + jsonFileName;
            File jsonFile = new File(jsonFilePath);
            if (jsonFile.exists()) {
                String excelFileName = jsonFileName.replace(".json", ".xlsx");
                String excelFilePath = basePath + File.separator + excelFileName;
                FileOutputStream outputStream = new FileOutputStream(excelFilePath);
                Workbook workbook = new XSSFWorkbook();

                ObjectMapper objectMapper = new ObjectMapper();
                JsonNode rootNode = objectMapper.readTree(jsonFile);

                // Iterate over each sheet in the JSON
                Iterator<Map.Entry<String, JsonNode>> fields = rootNode.fields();
                while (fields.hasNext()) {
                    Map.Entry<String, JsonNode> entry = fields.next();
                    String sheetName = entry.getKey();
                    JsonNode jsonArray = entry.getValue();

                    Sheet sheet = workbook.createSheet(sheetName);
                    int rowNum = 0;

                    // Write the header row with object names
                    Row headerRow = sheet.createRow(rowNum++);
                    int colNum = 0;
                    for (Iterator<String> fieldNames = jsonArray.get(0).fieldNames(); fieldNames.hasNext(); ) {
                        String fieldName = fieldNames.next();
                        Cell headerCell = headerRow.createCell(colNum++);
                        headerCell.setCellValue(fieldName);
                    }

                    // Iterate over each object in the array and create a row in the sheet
                    for (JsonNode row : jsonArray) {
                        Row excelRow = sheet.createRow(rowNum++);
                        colNum = 0;
                        for (Iterator<String> fieldNames = row.fieldNames(); fieldNames.hasNext(); ) {
                            String fieldName = fieldNames.next();
                            JsonNode fieldValue = row.get(fieldName);
                            Cell cell = excelRow.createCell(colNum++);
                            cell.setCellValue(fieldValue.asText());
                        }
                    }
                }

                workbook.write(outputStream);
                workbook.close();
                outputStream.close();

                System.out.println("XLSX file has been created successfully!");
            } else {
                System.out.println("No JSON file found at: " + jsonFilePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }




    private static String findFileName(String directory, String extension) {
        File dir = new File(directory);
        File[] files = dir.listFiles((dir1, name) -> name.toLowerCase().endsWith(extension));
        if (files != null && files.length > 0) {
            return files[0].getName(); // Return just the filename
        } else {
            return null;
        }
    }

    private static String cellToString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}

