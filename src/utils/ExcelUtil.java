package utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.xssf.usermodel.*;
import org.dom4j.DocumentException;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtil {
    private static final String KEY_FLAG = "key";
    public static final String ANNOTATION_FLAG = "##ANN##";
    public static final String DEFAULT_FLAG = "default";

    public void readExcel(File inputFile, File outputFile) throws IOException {
        if (!inputFile.exists() || !outputFile.exists()) {
            throw new IOException("文件不存在");
        }
        String fileName = inputFile.getName();
        String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName.substring(fileName.lastIndexOf(".") + 1);
        if ("xls".equals(extension)) {
            readXLSExcel(inputFile, outputFile);
        } else if ("xlsx".equals(extension)) {
            readXLSXExcel(inputFile, outputFile);
        } else {
            throw new IOException("不支持的文件类型");
        }
    }

    public void writXLSXExcel(File inputFile, File outputFile) throws IOException {
        if (!inputFile.exists() || !outputFile.exists()) {
            throw new IOException("文件不存在");
        }
        try {
            Map<String, Object> map = XMLUtil.readFormatXML(inputFile);
            XSSFWorkbook xwb = new XSSFWorkbook();
            XSSFSheet sheet = xwb.createSheet("Sheet1");
            //第三步创建行row:添加表头0行
            XSSFRow row = sheet.createRow(0);
            row.createCell(0).setCellValue(KEY_FLAG);
            row.createCell(1).setCellValue(DEFAULT_FLAG);
            XSSFCellStyle style = xwb.createCellStyle();
            int index = 1;
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                XSSFRow createRow = sheet.createRow(index);
                index = index + 1;
                if (entry.getKey().contains(ANNOTATION_FLAG)) {
                    createRow.createCell(0).setCellValue("<!-- " + entry.getValue().toString() + " -->");
                    continue;
                }
                createRow.createCell(0).setCellValue(entry.getKey());
                createRow.createCell(1).setCellValue(entry.getValue().toString());
            }
            //将excel写入
            OutputStream stream = new FileOutputStream(outputFile.getAbsolutePath() + File.separator + getFileNameNoEx(inputFile.getName()) + ".xlsx");
            xwb.write(stream);
            stream.close();
        } catch (DocumentException e) {
            e.printStackTrace();
        }
    }

    public static void writXLSXExcel(String mapName, int columnIndex, Map<String, Object> map, String outputFilePath) {
        File outputFile = new File(outputFilePath);
        if (!outputFile.exists()) {
            try {
                outputFile.createNewFile();
                XSSFWorkbook xwb = new XSSFWorkbook();
                FileOutputStream outputStream = new FileOutputStream(outputFile);
                xwb.write(outputStream);
                outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println(columnIndex + mapName);

        XSSFWorkbook xwb = null;
        try {
            xwb = new XSSFWorkbook(new FileInputStream(outputFile));
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            XSSFSheet sheet = xwb.getSheet("Sheet1");

            if ("default".equalsIgnoreCase(mapName) && sheet == null) {
                System.out.println("开始创建Sheet1");
                sheet = xwb.createSheet("Sheet1");
                //第三步创建行row:添加表头0行
                XSSFRow row = sheet.createRow(0);
                row.createCell(0).setCellValue(KEY_FLAG);
                row.createCell(1).setCellValue(DEFAULT_FLAG);

                int index = 1;
                for (Map.Entry<String, Object> entry : map.entrySet()) {
                    XSSFRow createRow = sheet.createRow(index);
                    index = index + 1;
                    if (entry.getKey().contains(ANNOTATION_FLAG)) {
                        createRow.createCell(0).setCellValue("<!-- " + entry.getValue().toString() + " -->");
                        continue;
                    }
                    createRow.createCell(0).setCellValue(entry.getKey());
                    createRow.createCell(1).setCellValue(entry.getValue().toString());
                }
            } else {
                System.out.println("开始修改");
                XSSFRow titleRow = sheet.getRow(0);
                titleRow.createCell(columnIndex).setCellValue(mapName);
                int keyIndex = 1;
                while (true) {
                    XSSFRow row = sheet.getRow(keyIndex++);
                    if (row == null) {
                        break;
                    }
                    XSSFCell cell = null;
                    cell = row.getCell(0);
                    if (cell == null) {
                        break;
                    }
                    String keyStr = cell.getStringCellValue();
                    String valueStr = (String) map.get(keyStr);
                    row.createCell(columnIndex).setCellValue(valueStr);
                }
            }

            //将excel写入

            OutputStream stream = new FileOutputStream(outputFilePath);
            xwb.write(stream);
            stream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void readXLSExcel(File file, File outputFile) {
        try {
            HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
            HSSFSheet sheet = hwb.getSheetAt(0);
            HSSFRow row = sheet.getRow(sheet.getFirstRowNum());
            if (row == null) {
                throw new IOException("row 不存在");
            }
            HSSFCell cell = row.getCell(row.getFirstCellNum());
            String keyCell = cell.getStringCellValue();
            //第一行 第一列必须为key字段
            if (!KEY_FLAG.equalsIgnoreCase(keyCell)) {
                throw new IOException("key 不存在");
            }
            int startIndex = row.getFirstCellNum() + 1;
            int endIndex = row.getLastCellNum();
            List<String> fileList = new ArrayList<>();
            for (int i = startIndex; i < endIndex; i++) {
                fileList.add(row.getCell(i).getStringCellValue());
            }
            //从第一列开始读
            for (int i = startIndex; i < endIndex; i++) {
                Map<String, Object> map = new LinkedHashMap<>();
                for (int j = sheet.getFirstRowNum(); j < sheet
                        .getPhysicalNumberOfRows(); j++) {
                    if (j == 0) {
                        //跳过第一行
                        continue;
                    }
                    row = sheet.getRow(j);
                    if (row == null) {
                        break;
                    }
                    String key = row.getCell(row.getFirstCellNum()).getStringCellValue();
                    if (containAnnotationKey(key)) {
                        map.put(ANNOTATION_FLAG + j, key.replace("<!--", "").replace("-->", ""));
                        continue;
                    }
                    HSSFCell currentCell = sheet.getRow(j).getCell(i);
                    Object value = getCellValue(currentCell);
                    map.put(key, value);
                }
                File xmlFile = new File(outputFile + File.separator + "string_" + fileList.get(i - 1) + ".xml");
                XMLUtil.writFormatXML(xmlFile, map);
            }
        } catch (OfficeXmlFileException officeXmlFileException) {
            readXLSXExcel(file, outputFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void readXLSXExcel(File file, File outputFile) {
        try {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
            // 读取第一章表格内容
            XSSFSheet sheet = xwb.getSheetAt(0);
            //获取第一行 第一列
            XSSFRow row = sheet.getRow(sheet.getFirstRowNum());
            if (row == null) {
                throw new IOException("row 不存在");
            }
            String keyCell = row.getCell(row.getFirstCellNum()).getStringCellValue();
            //第一行 第一列必须为key字段
            if (!KEY_FLAG.equalsIgnoreCase(keyCell)) {
                throw new IOException("key 不存在");
            }
            int startIndex = row.getFirstCellNum() + 1;
            int endIndex = row.getLastCellNum();
            //获取Top文件目录
            List<String> fileList = new ArrayList<>();
            for (int i = startIndex; i < endIndex; i++) {
                fileList.add(row.getCell(i).getStringCellValue());
            }
            //从第一列开始读
            for (int i = startIndex; i < endIndex; i++) {
                Map<String, Object> map = new LinkedHashMap<>();
                for (int j = sheet.getFirstRowNum(); j < sheet
                        .getPhysicalNumberOfRows(); j++) {
                    if (j == 0) {
                        //跳过第一行
                        continue;
                    }
                    row = sheet.getRow(j);
                    if (row == null) {
                        break;
                    }

                    String key = row.getCell(row.getFirstCellNum()).getStringCellValue();
                    if (containAnnotationKey(key)) {
                        map.put(ANNOTATION_FLAG + j, key.replace("<!--", "").replace("-->", ""));
                        continue;
                    }
                    XSSFCell currentCell = sheet.getRow(j).getCell(i);
                    Object value = getCellValue(currentCell);
                    map.put(key, value);
                }
                File xmlFile = new File(outputFile + File.separator + "string_" + fileList.get(i - 1) + ".xml");
                System.out.println(map.toString());
                XMLUtil.writFormatXML(xmlFile, map);
            }
        } catch (OfficeXmlFileException officeXmlFileException) {
            readXLSExcel(file, outputFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static Map<String, Map<String, Object>> readExcel(String inputPath, String outputDirPath) {
        try {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(inputPath));
            // 读取第一章表格内容
            XSSFSheet sheet = xwb.getSheetAt(0);
            //获取第一行 第一列
            XSSFRow row = sheet.getRow(sheet.getFirstRowNum());
            if (row == null) {
                throw new IOException("row 不存在");
            }
            String keyCell = row.getCell(row.getFirstCellNum()).getStringCellValue();
            //第一行 第一列必须为key字段
            if (!KEY_FLAG.equalsIgnoreCase(keyCell)) {
                throw new IOException("key 不存在");
            }
            int startIndex = row.getFirstCellNum() + 1;
            int endIndex = row.getLastCellNum();
            //获取Top文件目录
            List<String> fileList = new ArrayList<>();
            for (int i = startIndex; i < endIndex; i++) {
                String titleStr = row.getCell(i).getStringCellValue();
                if ("default".equalsIgnoreCase(titleStr)) {
                    titleStr = "values";
                } else {
                    titleStr = "values" + titleStr;
                }
                fileList.add(titleStr);
            }
            //从第一列开始读
            for (int i = startIndex; i < endIndex; i++) {
                Map<String, Object> map = new LinkedHashMap<>();
                for (int j = sheet.getFirstRowNum(); j < sheet.getPhysicalNumberOfRows(); j++) {
                    if (j == 0) {
                        //跳过第一行
                        continue;
                    }
                    row = sheet.getRow(j);
                    if (row == null) {
                        break;
                    }

                    String key = row.getCell(row.getFirstCellNum()).getStringCellValue();
                    if (containAnnotationKey(key)) {
                        map.put(ANNOTATION_FLAG + j, key.replace("<!--", "").replace("-->", ""));
                        continue;
                    }
                    XSSFCell currentCell = sheet.getRow(j).getCell(i);
                    Object value = getCellValue(currentCell);
                    map.put(key, value);
                }
                File outputDir = (new File(outputDirPath + File.separator + fileList.get(i - 1)));
                if (!outputDir.exists()) {
                    outputDir.mkdirs();
                }
                File xmlFile = new File(outputDir.getAbsolutePath() + File.separator + "strings.xml");
                if (!xmlFile.exists()) {
                    xmlFile.createNewFile();
                }
                System.out.println(xmlFile.getAbsolutePath());

                XMLUtil.writFormatXML(xmlFile, map);
            }
        } catch (OfficeXmlFileException officeXmlFileException) {
//            readXLSExcel(file, outputFile);
        } catch (IOException e) {
            e.printStackTrace();
        }

        Map<String, Map<String, Object>> maps = new HashMap<>();

        return maps;
    }


    private void makeDirectory(File parent, String child) {
        File file = new File(parent, child);
        if (file.exists()) {
            // 文件已经存在，输出文件的相关信息
            System.out.println(file.getAbsolutePath());
            System.out.println(file.getName());
            System.out.println(file.length());
        } else {
            file.getParentFile().mkdirs();
        }
    }

    private static boolean containAnnotationKey(String key) {
        return key.contains("<!--") && key.contains("-->");
    }

    private Object getCellValue(HSSFCell cell) {
        Object value;
        if (cell == null) {
            return value = "";
        }
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                DecimalFormat df = new DecimalFormat("0");// 格式化 number String
                // 字符
                SimpleDateFormat sdf = new SimpleDateFormat(
                        "yyyy-MM-dd");// 格式化日期字符串
                DecimalFormat nf = new DecimalFormat("0");// 格式化数字

                if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                    value = df.format(cell.getNumericCellValue());
                } else if ("General".equals(cell.getCellStyle()
                        .getDataFormatString())) {
                    value = nf.format(cell.getNumericCellValue());
                } else {
                    value = sdf.format(HSSFDateUtil.getJavaDate(cell
                            .getNumericCellValue()));
                }
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                value = "";
                break;
            default:
                value = cell.toString();
        }

        return value;
    }

    private static Object getCellValue(XSSFCell cell) {
        Object value;
        if (cell == null) {
            return value = "";
        }
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                DecimalFormat df = new DecimalFormat("0");// 格式化 number String
                // 字符
                SimpleDateFormat sdf = new SimpleDateFormat(
                        "yyyy-MM-dd");// 格式化日期字符串
                DecimalFormat nf = new DecimalFormat("0");// 格式化数字

                if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                    value = df.format(cell.getNumericCellValue());
                } else if ("General".equals(cell.getCellStyle()
                        .getDataFormatString())) {
                    value = nf.format(cell.getNumericCellValue());
                } else {
                    value = sdf.format(HSSFDateUtil.getJavaDate(cell
                            .getNumericCellValue()));
                }
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                value = "";
                break;
            default:
                value = cell.toString();
        }
        return value;
    }

    public static String getFileNameNoEx(String filename) {
        if ((filename != null) && (filename.length() > 0)) {
            int dot = filename.lastIndexOf('.');
            if ((dot > -1) && (dot < (filename.length()))) {
                return filename.substring(0, dot);
            }
        }
        return filename;
    }
}
