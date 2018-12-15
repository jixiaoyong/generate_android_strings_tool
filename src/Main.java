import org.dom4j.DocumentException;
import utils.ExcelUtil;
import utils.XMLUtil;

import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

public class Main {

    static String projectName = "{project_name}";
    static String resDirPath = "E:\\jixiaoyong\\learn\\AndroidTranslationTools\\res\\";
    static String intputDirPath = "E:\\jixiaoyong\\learn\\AndroidTranslationTools\\res\\input\\";
    static String inputXmlDirPath = resDirPath + "input\\" + projectName + "\\";
    static String outputExcelPath = resDirPath + "output\\excel\\" + projectName + ".xlsx";
    static String outputXmlDirPath = resDirPath + "output\\xml\\" + projectName + "\\";

    public static void main(String[] args) {
        System.out.println("Hello World!");
        File projectDir = new File(intputDirPath);
        File[] projects = projectDir.listFiles();

            String projectNameStr = "project";

            writeToExcel(outputExcelPath.replace(projectName, projectNameStr),
                    readXmls(inputXmlDirPath.replace(projectName, projectNameStr)));

//        writeToExcel(outputExcelPath, readXmls(inputXmlDirPath));
//        ExcelUtil.readExcel(outputExcelPath, outputXmlDirPath);
    }


    public static Map<String, Map<String, Object>> readXmls(String path) {
        Map<String, Map<String, Object>> maps = new HashMap<>();
        File file = new File(path);
        File[] xmlFiles = file.listFiles();
        for (int i = 0; i < xmlFiles.length; i++) {
            if (!xmlFiles[i].getName().startsWith("values")) {
                continue;
            }
            try {
                String language = xmlFiles[i].getAbsolutePath().substring(xmlFiles[i].getAbsolutePath().lastIndexOf("\\") + 1)
                        .replace("values", "");
                if ("".equals(language)) {
                    language = "default";
                }
                Map<String, Object> map = null;
                try {
                    map = XMLUtil.readFormatXML(new File(xmlFiles[i].getAbsolutePath() + "/strings.xml"));
                } catch (DocumentException e) {
                    map = XMLUtil.readFormatXML(new File(xmlFiles[i].getAbsolutePath() + "/string.xml"));
                    e.printStackTrace();
                }
                maps.put(language, map);
            } catch (DocumentException e) {
                e.printStackTrace();
            }
        }
        return maps;
    }

    public static Map<String, Map<String, Object>> readExcel(String path) {
        Map<String, Map<String, Object>> maps = new HashMap<>();


        return maps;
    }

    public static void writeToExcel(String path, Map<String, Map<String, Object>> maps) {
        Set<String> mapKetSet = maps.keySet();
        ExcelUtil.writXLSXExcel("default", 1, maps.get("default"), path);
        maps.remove("default");
        for (int i = 0; i < mapKetSet.size(); i++) {
            ExcelUtil.writXLSXExcel((String) mapKetSet.toArray()[i], i + 2, maps.get(mapKetSet.toArray()[i]), path);
        }
    }

    public static void writeToXML(String path, Map<String, Map<String, Object>> maps) {
        Set<String> mapKetSet = maps.keySet();

        ExcelUtil.writXLSXExcel("default", 1, maps.get("default"), path);
        maps.remove("default");
        for (int i = 0; i < mapKetSet.size(); i++) {
            ExcelUtil.writXLSXExcel((String) mapKetSet.toArray()[i], i + 2, maps.get(mapKetSet.toArray()[i]), path);
        }
    }
}